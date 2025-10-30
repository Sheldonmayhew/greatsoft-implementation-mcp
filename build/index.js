import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema, } from "@modelcontextprotocol/sdk/types.js";
import sql from "mssql";
import * as fs from "fs/promises";
import * as XLSX from "xlsx";
import { v4 as uuidv4 } from "uuid";
class GreatSoftImplementation {
    config;
    pool = null;
    countryID = 1;
    constructor(config) {
        this.config = {
            ...config,
            options: {
                encrypt: true,
                trustServerCertificate: false,
                ...config.options,
            },
        };
    }
    async connect() {
        if (!this.pool) {
            this.pool = await sql.connect(this.config);
        }
    }
    async disconnect() {
        if (this.pool) {
            await this.pool.close();
            this.pool = null;
        }
    }
    setCountryID(countryID) {
        this.countryID = countryID;
    }
    async licenseDatabase(scriptPath) {
        await this.connect();
        try {
            const script = await fs.readFile(scriptPath, "utf-8");
            const batches = script
                .split(/^\s*GO\s*$/gim)
                .map((batch) => batch.trim())
                .filter((batch) => batch.length > 0);
            for (const batch of batches) {
                await this.pool.request().query(batch);
            }
            return {
                success: true,
                recordsImported: 0,
                errors: [],
                message: `Database licensed successfully. Executed ${batches.length} SQL batches.`,
            };
        }
        catch (error) {
            return {
                success: false,
                recordsImported: 0,
                errors: [{ row: 0, field: "script", value: "", message: error.message }],
                message: `Failed to license database: ${error.message}`,
            };
        }
    }
    async readExcelFile(filePath) {
        const buffer = await fs.readFile(filePath);
        const workbook = XLSX.read(buffer, { cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, {
            range: 1,
            defval: null
        });
        return data;
    }
    mapExcelToSQL(excelRow, mappings) {
        const sqlRow = {};
        for (const [excelCol, sqlCol] of Object.entries(mappings)) {
            let value = excelRow[excelCol];
            if (value === "NULL" || value === "null") {
                value = null;
            }
            sqlRow[sqlCol] = value;
        }
        return sqlRow;
    }
    validateOfficeData(data, rowOffset = 2) {
        const errors = [];
        const seenCodes = new Set();
        data.forEach((row, index) => {
            const actualRow = index + rowOffset;
            if (row.OfficeCode === "OfficeCode" || row.OfficeCode === "Data" || row.OfficeCode === "Context") {
                return;
            }
            if (!row.OfficeCode) {
                errors.push({
                    row: actualRow,
                    field: "OfficeCode",
                    value: row.OfficeCode,
                    message: "Office Code is required and will be auto-generated if missing",
                });
            }
            if (!row.OfficeDesc) {
                errors.push({
                    row: actualRow,
                    field: "OfficeDesc",
                    value: row.OfficeDesc,
                    message: "Office Name is required",
                });
            }
            if (row.OfficeCode && seenCodes.has(row.OfficeCode)) {
                errors.push({
                    row: actualRow,
                    field: "OfficeCode",
                    value: row.OfficeCode,
                    message: `Duplicate Office Code: ${row.OfficeCode}`,
                });
            }
            if (row.OfficeCode) {
                seenCodes.add(row.OfficeCode);
            }
            if (row.OfficeCode && row.OfficeCode.length > 10) {
                errors.push({
                    row: actualRow,
                    field: "OfficeCode",
                    value: row.OfficeCode,
                    message: "Office Code must be 10 characters or less",
                });
            }
            if (row.OfficeDesc && row.OfficeDesc.length > 100) {
                errors.push({
                    row: actualRow,
                    field: "OfficeDesc",
                    value: row.OfficeDesc,
                    message: "Office Name must be 100 characters or less",
                });
            }
        });
        return errors;
    }
    async importOffices(excelPath) {
        await this.connect();
        try {
            const rawData = await this.readExcelFile(excelPath);
            const columnMap = {
                "OfficeCode": "OfficeCode",
                "OfficeDesc": "OfficeDesc",
                "OfficePAddress": "OfficePAddress",
                "OfficeBussAdd": "OfficeBussAdd",
                "OfficeTel": "OfficeTel",
                "OfficeFax": "OfficeFax",
                "OfficeBank": "OfficeBank",
                "OfficeBranch": "OfficeBranch",
                "OfficeBranchNo": "OfficeBranchNo",
                "OfficeBankAcc": "OfficeBankAcc",
                "OfficeRegNo": "OfficeRegNo",
                "OfficeTaxNo": "OfficeTaxNo",
                "OfficeURL": "OfficeURL",
                "OfficeEmail": "OfficeEmail",
            };
            const offices = rawData
                .map(row => this.mapExcelToSQL(row, columnMap))
                .filter(row => row.OfficeDesc);
            const validationErrors = this.validateOfficeData(offices);
            if (validationErrors.length > 0) {
                const criticalErrors = validationErrors.filter(e => e.message.includes("required") || e.message.includes("Duplicate"));
                if (criticalErrors.length > 0) {
                    return {
                        success: false,
                        recordsImported: 0,
                        errors: validationErrors,
                        message: `Validation failed with ${criticalErrors.length} critical error(s)`,
                    };
                }
            }
            const generatedCodes = {};
            for (const office of offices) {
                if (!office.OfficeCode) {
                    const baseCode = office.OfficeDesc
                        .substring(0, 10)
                        .toUpperCase()
                        .replace(/[^A-Z0-9]/g, "");
                    let code = baseCode;
                    let counter = 1;
                    const existingCodes = offices.map(o => o.OfficeCode).filter(Boolean);
                    while (existingCodes.includes(code)) {
                        code = `${baseCode.substring(0, 8)}${counter.toString().padStart(2, "0")}`;
                        counter++;
                    }
                    office.OfficeCode = code;
                    generatedCodes[office.OfficeDesc] = code;
                }
            }
            let recordsImported = 0;
            for (const office of offices) {
                try {
                    const officeID = uuidv4();
                    const request = this.pool.request();
                    request.input("CountryID", sql.Int, this.countryID);
                    request.input("OfficeCode", sql.NVarChar(10), office.OfficeCode);
                    request.input("OfficeID", sql.UniqueIdentifier, officeID);
                    request.input("OfficeDesc", sql.NVarChar(100), office.OfficeDesc);
                    request.input("OfficePAddress", sql.NVarChar(255), office.OfficePAddress);
                    request.input("OfficeBussAdd", sql.NVarChar(255), office.OfficeBussAdd);
                    request.input("OfficeTel", sql.NVarChar(30), office.OfficeTel);
                    request.input("OfficeFax", sql.NVarChar(30), office.OfficeFax);
                    request.input("OfficeBank", sql.NVarChar(50), office.OfficeBank);
                    request.input("OfficeBranch", sql.NVarChar(50), office.OfficeBranch);
                    request.input("OfficeBranchNo", sql.NVarChar(20), office.OfficeBranchNo);
                    request.input("OfficeBankAcc", sql.NVarChar(50), office.OfficeBankAcc);
                    request.input("OfficeRegNo", sql.NVarChar(30), office.OfficeRegNo);
                    request.input("OfficeTaxNo", sql.NVarChar(30), office.OfficeTaxNo);
                    request.input("OfficeURL", sql.NVarChar(255), office.OfficeURL);
                    request.input("OfficeEmail", sql.NVarChar(100), office.OfficeEmail);
                    await request.query(`
            INSERT INTO dbo.Office (
              CountryID, OfficeCode, OfficeID, OfficeDesc,
              OfficePAddress, OfficeBussAdd, OfficeTel, OfficeFax,
              OfficeBank, OfficeBranch, OfficeBranchNo, OfficeBankAcc,
              OfficeRegNo, OfficeTaxNo, OfficeURL, OfficeEmail
            ) VALUES (
              @CountryID, @OfficeCode, @OfficeID, @OfficeDesc,
              @OfficePAddress, @OfficeBussAdd, @OfficeTel, @OfficeFax,
              @OfficeBank, @OfficeBranch, @OfficeBranchNo, @OfficeBankAcc,
              @OfficeRegNo, @OfficeTaxNo, @OfficeURL, @OfficeEmail
            )
          `);
                    recordsImported++;
                }
                catch (error) {
                    validationErrors.push({
                        row: offices.indexOf(office) + 2,
                        field: "database",
                        value: office.OfficeCode,
                        message: `Database error: ${error.message}`,
                    });
                }
            }
            return {
                success: recordsImported > 0,
                recordsImported,
                errors: validationErrors,
                message: `Successfully imported ${recordsImported} of ${offices.length} office(s)`,
                generatedCodes: Object.keys(generatedCodes).length > 0 ? generatedCodes : undefined,
            };
        }
        catch (error) {
            return {
                success: false,
                recordsImported: 0,
                errors: [{ row: 0, field: "general", value: "", message: error.message }],
                message: `Failed to import offices: ${error.message}`,
            };
        }
    }
    async getImplementationStatus() {
        await this.connect();
        const officeCount = await this.pool
            .request()
            .query("SELECT COUNT(*) as count FROM dbo.Office");
        const employeeCount = await this.pool
            .request()
            .query("SELECT COUNT(*) as count FROM dbo.Employee");
        const clientCount = await this.pool
            .request()
            .query("SELECT COUNT(*) as count FROM dbo.Client");
        return {
            offices: officeCount.recordset[0].count,
            employees: employeeCount.recordset[0].count,
            clients: clientCount.recordset[0].count,
            status: "In Progress",
        };
    }
}
const server = new Server({
    name: "greatsoft-implementation",
    version: "1.0.0",
}, {
    capabilities: {
        tools: {},
    },
});
let implementation = null;
server.setRequestHandler(ListToolsRequestSchema, async () => {
    return {
        tools: [
            {
                name: "configure_database",
                description: "Configure the MS SQL Server connection for GreatSoft implementation. Must be called first.",
                inputSchema: {
                    type: "object",
                    properties: {
                        server: {
                            type: "string",
                            description: "SQL Server hostname or IP",
                        },
                        database: {
                            type: "string",
                            description: "Database name",
                        },
                        user: {
                            type: "string",
                            description: "SQL Server username",
                        },
                        password: {
                            type: "string",
                            description: "SQL Server password",
                        },
                        port: {
                            type: "number",
                            description: "SQL Server port (default: 1433)",
                        },
                        countryID: {
                            type: "number",
                            description: "Country ID for offices (default: 1)",
                        },
                    },
                    required: ["server", "database", "user", "password"],
                },
            },
            {
                name: "license_database",
                description: "Run the SQL licensing script to prepare the database for client data import.",
                inputSchema: {
                    type: "object",
                    properties: {
                        scriptPath: {
                            type: "string",
                            description: "Path to the SQL licensing script file",
                        },
                    },
                    required: ["scriptPath"],
                },
            },
            {
                name: "import_offices",
                description: "Import office data from the client's completed Excel file. Validates data and auto-generates office codes if missing.",
                inputSchema: {
                    type: "object",
                    properties: {
                        excelPath: {
                            type: "string",
                            description: "Path to the GreatSoft Office Excel file",
                        },
                    },
                    required: ["excelPath"],
                },
            },
            {
                name: "get_implementation_status",
                description: "Get current implementation status showing counts of offices, employees, and clients imported.",
                inputSchema: {
                    type: "object",
                    properties: {},
                },
            },
        ],
    };
});
server.setRequestHandler(CallToolRequestSchema, async (request) => {
    try {
        if (request.params.name === "configure_database") {
            const args = request.params.arguments;
            const config = {
                server: args.server,
                database: args.database,
                user: args.user,
                password: args.password,
                port: args.port || 1433,
                options: {
                    encrypt: true,
                    trustServerCertificate: true,
                },
            };
            implementation = new GreatSoftImplementation(config);
            if (args.countryID) {
                implementation.setCountryID(args.countryID);
            }
            await implementation.connect();
            return {
                content: [
                    {
                        type: "text",
                        text: `✓ Connected to SQL Server: ${config.server}/${config.database}\n✓ Country ID set to: ${args.countryID || 1}\n\nReady for implementation!`,
                    },
                ],
            };
        }
        if (!implementation) {
            throw new Error("Database not configured. Call configure_database first.");
        }
        if (request.params.name === "license_database") {
            const args = request.params.arguments;
            const result = await implementation.licenseDatabase(args.scriptPath);
            return {
                content: [
                    {
                        type: "text",
                        text: JSON.stringify(result, null, 2),
                    },
                ],
            };
        }
        if (request.params.name === "import_offices") {
            const args = request.params.arguments;
            const result = await implementation.importOffices(args.excelPath);
            let message = result.message;
            if (result.generatedCodes && Object.keys(result.generatedCodes).length > 0) {
                message += "\n\n**Auto-generated Office Codes:**\n";
                for (const [name, code] of Object.entries(result.generatedCodes)) {
                    message += `  • ${name}: ${code}\n`;
                }
            }
            if (result.errors.length > 0) {
                message += "\n\n**Validation Warnings:**\n";
                result.errors.forEach((error) => {
                    message += `  • Row ${error.row}, ${error.field}: ${error.message}\n`;
                });
            }
            return {
                content: [
                    {
                        type: "text",
                        text: message,
                    },
                ],
            };
        }
        if (request.params.name === "get_implementation_status") {
            const status = await implementation.getImplementationStatus();
            return {
                content: [
                    {
                        type: "text",
                        text: `**GreatSoft Implementation Status**\n\n` +
                            `Offices: ${status.offices}\n` +
                            `Employees: ${status.employees}\n` +
                            `Clients: ${status.clients}\n\n` +
                            `Status: ${status.status}`,
                    },
                ],
            };
        }
        throw new Error(`Unknown tool: ${request.params.name}`);
    }
    catch (error) {
        return {
            content: [
                {
                    type: "text",
                    text: `❌ Error: ${error.message}`,
                },
            ],
            isError: true,
        };
    }
});
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("GreatSoft Implementation MCP Server running");
}
main().catch((error) => {
    console.error("Fatal error:", error);
    process.exit(1);
});
