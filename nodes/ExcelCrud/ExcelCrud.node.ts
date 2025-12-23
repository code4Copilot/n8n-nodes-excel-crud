import {
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeExecutionData,
	INodePropertyOptions,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';

import * as ExcelJS from 'exceljs';
import * as fs from 'fs/promises';
import * as path from 'path';

export class ExcelCrud implements INodeType {
	methods = {
		loadOptions: {
			async getWorksheets(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const inputMode = this.getNodeParameter('inputMode') as string;
				if (inputMode === 'filePath') {
					const filePath = this.getNodeParameter('filePath') as string;
					
					if (!filePath) {
						// 如果沒有提供檔案路徑，返回空陣列
						return [];
					}
					
					const workbook = new ExcelJS.Workbook();
					try {
						await fs.access(filePath);
						await workbook.xlsx.readFile(filePath);
						
						const worksheets = workbook.worksheets.map(ws => ({
							name: ws.name,
							value: ws.name,
						}));
						
						if (worksheets.length === 0) {
							return [{ name: 'No worksheets found', value: '' }];
						}
						
						return worksheets;
					} catch (error) {
						// 返回錯誤訊息作為選項
						const errorMessage = error instanceof Error ? error.message : 'Unknown error';
						return [{ 
							name: `Error: Cannot read file - ${errorMessage}`, 
							value: '' 
						}];
					}
				}
				return [];
			},
			async getColumns(this: ILoadOptionsFunctions): Promise<INodePropertyOptions[]> {
				const inputMode = this.getNodeParameter('inputMode') as string;
				let sheetName = '';
				
				if (inputMode === 'filePath') {
					const filePath = this.getNodeParameter('filePath') as string;
					sheetName = this.getNodeParameter('sheetNameOptions', '') as string;
					
					if (!filePath || !sheetName) {
						return [];
					}
					
					const workbook = new ExcelJS.Workbook();
					try {
						await fs.access(filePath);
						await workbook.xlsx.readFile(filePath);
						
						const worksheet = workbook.getWorksheet(sheetName);
						if (!worksheet) {
							return [{ name: 'Worksheet not found', value: '' }];
						}
						
						const columns: INodePropertyOptions[] = [];
					worksheet.getRow(1).eachCell((cell, _colNumber) => {
							const columnName = cell.value as string;
							if (columnName) {
								columns.push({
									name: columnName,
									value: columnName,
								});
							}
						});
						
						if (columns.length === 0) {
							return [{ name: 'No columns found', value: '' }];
						}
						
						return columns;
					} catch (error) {
						const errorMessage = error instanceof Error ? error.message : 'Unknown error';
						return [{ 
							name: `Error: Cannot read file - ${errorMessage}`, 
							value: '' 
						}];
					}
				} else if (inputMode === 'binaryData') {
					// Binary Data 模式下無法動態載入，返回提示
					return [{ name: 'Please use file path mode for column selection', value: '' }];
				}
				return [];
			},
		},
	};

	description: INodeTypeDescription = {
		displayName: 'Excel CRUD',
		name: 'excelCrud',
		icon: 'file:excel-crud.svg',
		group: ['transform'],
		version: 1,
		description: 'Perform CRUD operations on Excel files with direct file access',
		defaults: {
			name: 'Excel CRUD',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			// ... (properties 部分保持不變，略)
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{ name: 'Row', value: 'row' },
					{ name: 'Worksheet', value: 'worksheet' },
				],
				default: 'row',
			},
			{
				displayName: 'Input Mode',
				name: 'inputMode',
				type: 'options',
				options: [
					{ name: 'File Path', value: 'filePath' },
					{ name: 'Binary Data', value: 'binaryData' },
				],
				default: 'filePath',
			},
			{
				displayName: 'File Path',
				name: 'filePath',
				type: 'string',
				displayOptions: { show: { inputMode: ['filePath'] } },
				default: '',
			},
			{
				displayName: 'Binary Property',
				name: 'binaryPropertyName',
				type: 'string',
				displayOptions: { show: { inputMode: ['binaryData'] } },
				default: 'data',
			},
			{
				displayName: 'Auto Save',
				name: 'autoSave',
				type: 'boolean',
				displayOptions: { show: { inputMode: ['filePath'], resource: ['row'] } },
				default: true,
			},
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				displayOptions: { show: { resource: ['row'] } },
				options: [
					{ name: 'Append Row', value: 'appendRow' },
					{ name: 'Read Rows', value: 'readRows' },
					{ name: 'Update Row', value: 'updateRow' },
					{ name: 'Delete Row', value: 'deleteRow' },
					{ name: 'Find Rows', value: 'findRows' },
					{ name: 'Insert Row', value: 'insertRow' },
				],
				default: 'appendRow',
			},
			{
				displayName: 'Operation',
				name: 'worksheetOperation',
				type: 'options',
				displayOptions: { show: { resource: ['worksheet'] } },
				options: [
					{ name: 'Create Worksheet', value: 'createWorksheet' },
					{ name: 'Delete Worksheet', value: 'deleteWorksheet' },
					{ name: 'Rename Worksheet', value: 'renameWorksheet' },
					{ name: 'List Worksheets', value: 'listWorksheets' },
					{ name: 'Copy Worksheet', value: 'copyWorksheet' },
					{ name: 'Get Worksheet Info', value: 'getWorksheetInfo' },
				],
				default: 'listWorksheets',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				displayOptions: { show: { resource: ['row'], inputMode: ['binaryData'] } },
				default: 'Sheet1',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetNameOptions',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWorksheets',
					loadOptionsDependsOn: ['filePath'],
				},
				displayOptions: { show: { resource: ['row'], inputMode: ['filePath'] } },
				default: '',
				description: 'Select the worksheet to operate on',
				required: true,
			},
			{
				displayName: 'Sheet Name',
				name: 'worksheetName',
				type: 'string',
				displayOptions: { show: { resource: ['worksheet'], worksheetOperation: ['deleteWorksheet', 'renameWorksheet', 'copyWorksheet', 'getWorksheetInfo'], inputMode: ['binaryData'] } },
				default: 'Sheet1',
			},
			{
				displayName: 'Sheet Name',
				name: 'worksheetNameOptions',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getWorksheets',
					loadOptionsDependsOn: ['filePath'],
				},
				displayOptions: { show: { resource: ['worksheet'], worksheetOperation: ['deleteWorksheet', 'renameWorksheet', 'copyWorksheet', 'getWorksheetInfo'], inputMode: ['filePath'] } },
				default: '',
				description: 'Select the worksheet to operate on',
				required: true,
			},
			{
				displayName: 'New Sheet Name',
				name: 'newSheetName',
				type: 'string',
				displayOptions: { show: { resource: ['worksheet'], worksheetOperation: ['createWorksheet', 'renameWorksheet', 'copyWorksheet'] } },
				default: '',
			},
			{
				displayName: 'Initial Data',
				name: 'initialData',
				type: 'json',
				displayOptions: { show: { resource: ['worksheet'], worksheetOperation: ['createWorksheet'] } },
				default: '[]',
			},
			{
				displayName: 'Include Hidden',
				name: 'includeHidden',
				type: 'boolean',
				displayOptions: { show: { resource: ['worksheet'], worksheetOperation: ['listWorksheets'] } },
				default: false,
			},
			{
				displayName: 'Row Data',
				name: 'rowData',
				type: 'json',
				displayOptions: { show: { resource: ['row'], operation: ['appendRow', 'insertRow'] } },
				default: '{}',
			},
			{
				displayName: 'Start Row',
				name: 'startRow',
				type: 'number',
				displayOptions: { show: { resource: ['row'], operation: ['readRows'] } },
				default: 2,
			},
			{
				displayName: 'End Row',
				name: 'endRow',
				type: 'number',
				displayOptions: { show: { resource: ['row'], operation: ['readRows'] } },
				default: 0,
			},
			{
				displayName: 'Row Number',
				name: 'rowNumber',
				type: 'number',
				displayOptions: { show: { resource: ['row'], operation: ['updateRow', 'deleteRow', 'insertRow'] } },
				default: 2,
			},
			{
				displayName: 'Updated Data',
				name: 'updatedData',
				type: 'json',
				displayOptions: { show: { resource: ['row'], operation: ['updateRow'] } },
				default: '{}',
			},
			{
				displayName: 'Search Column',
				name: 'searchColumn',
				type: 'options',
				typeOptions: {
					loadOptionsMethod: 'getColumns',
					loadOptionsDependsOn: ['filePath', 'sheetNameOptions'],
				},
				displayOptions: { show: { resource: ['row'], operation: ['findRows'], inputMode: ['filePath'] } },
				default: '',
				description: 'Select the column to search in',
				required: true,
			},
			{
				displayName: 'Search Column',
				name: 'searchColumnManual',
				type: 'string',
				displayOptions: { show: { resource: ['row'], operation: ['findRows'], inputMode: ['binaryData'] } },
				default: '',
				description: 'Enter the column name or letter (e.g., "Name" or "A")',
				required: true,
			},
			{
				displayName: 'Search Value',
				name: 'searchValue',
				type: 'string',
				displayOptions: { show: { resource: ['row'], operation: ['findRows'] } },
				default: '',
			},
			{
				displayName: 'Match Type',
				name: 'matchType',
				type: 'options',
				displayOptions: { show: { resource: ['row'], operation: ['findRows'] } },
				options: [
					{ name: 'Exact Match', value: 'exact' },
					{ name: 'Contains', value: 'contains' },
					{ name: 'Starts With', value: 'startsWith' },
				],
				default: 'exact',
			},
			{
				displayName: 'Return Row Numbers',
				name: 'returnRowNumbers',
				type: 'boolean',
				displayOptions: { show: { resource: ['row'], operation: ['findRows'] } },
				default: false,
				description: 'Whether to return an array of row numbers instead of row data',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		for (let i = 0; i < items.length; i++) {
			try {
				const resource = this.getNodeParameter('resource', i) as string;
				const inputMode = this.getNodeParameter('inputMode', i) as string;

				const workbook = new ExcelJS.Workbook();
				let filePath: string | undefined;

				if (inputMode === 'filePath') {
					filePath = this.getNodeParameter('filePath', i) as string;
					try {
						await fs.access(filePath);
						await workbook.xlsx.readFile(filePath);
					} catch (error) {
						throw new NodeOperationError(this.getNode(), `無法讀取檔案: ${filePath}`, { itemIndex: i });
					}
				} else {
					const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
					this.helpers.assertBinaryData(i, binaryPropertyName);
					const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
					await workbook.xlsx.load(buffer as any);
				}

				let result: any;
			let operation: string = '';

			// 【關鍵修正】：使用類別名稱呼叫靜態方法
			if (resource === 'row') {
				operation = this.getNodeParameter('operation', i) as string;
				result = await ExcelCrud.executeRowOperation(this, workbook, i);
			} else if (resource === 'worksheet') {
				operation = this.getNodeParameter('worksheetOperation', i) as string;
				result = await ExcelCrud.executeWorksheetOperation(this, workbook, i);
			}

			if (inputMode === 'filePath') {
				const autoSave = this.getNodeParameter('autoSave', i, true) as boolean;
			if (resource === 'row' && ['readRows', 'findRows'].includes(operation)) {
				// 如果是 findRows 且返回 rowNumbers，則返回單一對象
				if (operation === 'findRows' && result.rowNumbers) {
					returnData.push({ json: result });
				} else {
					// readRows 和 findRows（一般模式）回傳陣列，每個項目作為一個 item
					for (const row of result) {
						returnData.push({ json: row });
					}
				}
				} else if (resource === 'worksheet' && ['listWorksheets', 'getWorksheetInfo'].includes(operation)) {
						returnData.push({ json: result });
					} else if (autoSave && filePath) {
						await workbook.xlsx.writeFile(filePath);
						returnData.push({ json: { ...result, filePath, saved: true } });
					} else {
						const outputBuffer = await workbook.xlsx.writeBuffer();
						const fileName = filePath ? path.basename(filePath) : 'output.xlsx';
						const newBinaryData = await this.helpers.prepareBinaryData((outputBuffer as unknown) as Buffer, fileName);
						returnData.push({ json: result, binary: { data: newBinaryData } });
					}
				} else {
					const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
					const binaryData = this.helpers.assertBinaryData(i, binaryPropertyName);
			if (resource === 'row' && ['readRows', 'findRows'].includes(operation)) {
				// 如果是 findRows 且返回 rowNumbers，則返回單一對象
				if (operation === 'findRows' && result.rowNumbers) {
					returnData.push({ json: result });
				} else {
					// readRows 和 findRows（一般模式）回傳陣列，每個項目作為一個 item
					for (const row of result) {
						returnData.push({ json: row });
					}
					}

				} else if (resource === 'worksheet' && ['listWorksheets', 'getWorksheetInfo'].includes(operation)) {
						returnData.push({ json: result });
					} else {
						const outputBuffer = await workbook.xlsx.writeBuffer();
						const newBinaryData = await this.helpers.prepareBinaryData((outputBuffer as unknown) as Buffer, binaryData.fileName || 'output.xlsx');
						returnData.push({ json: result || { success: true }, binary: { [binaryPropertyName]: newBinaryData } });
					}
				}
			} catch (error) {
				if (this.continueOnFail()) {
					returnData.push({ json: { error: (error as Error).message }, pairedItem: { item: i } });
					continue;
				}
				throw error;
			}
		}
		return [returnData];
	}

	// 所有方法改為 static
	private static async executeRowOperation(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number): Promise<any> {
		const operation = ef.getNodeParameter('operation', itemIndex) as string;
		const sheetName = ExcelCrud.getResolvedSheetName(ef, itemIndex, 'sheetName');
		const worksheet = workbook.getWorksheet(sheetName);
		if (!worksheet) throw new NodeOperationError(ef.getNode(), `工作表 "${sheetName}" 不存在`);

		switch (operation) {
			case 'appendRow': return await ExcelCrud.appendRow(ef, worksheet, itemIndex);
			case 'readRows': return await ExcelCrud.readRows(ef, worksheet, itemIndex);
			case 'updateRow': return await ExcelCrud.updateRow(ef, worksheet, itemIndex);
			case 'deleteRow': return await ExcelCrud.deleteRow(ef, worksheet, itemIndex);
			case 'findRows': return await ExcelCrud.findRows(ef, worksheet, itemIndex);
			case 'insertRow': return await ExcelCrud.insertRow(ef, worksheet, itemIndex);
			default: throw new NodeOperationError(ef.getNode(), `未知的操作: ${operation}`);
		}
	}

	private static async executeWorksheetOperation(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number): Promise<any> {
		const operation = ef.getNodeParameter('worksheetOperation', itemIndex) as string;
		switch (operation) {
			case 'createWorksheet': return await ExcelCrud.createWorksheet(ef, workbook, itemIndex);
			case 'deleteWorksheet': return await ExcelCrud.deleteWorksheet(ef, workbook, itemIndex);
			case 'renameWorksheet': return await ExcelCrud.renameWorksheet(ef, workbook, itemIndex);
			case 'listWorksheets': return await ExcelCrud.listWorksheets(ef, workbook, itemIndex);
			case 'copyWorksheet': return await ExcelCrud.copyWorksheet(ef, workbook, itemIndex);
			case 'getWorksheetInfo': return await ExcelCrud.getWorksheetInfo(ef, workbook, itemIndex);
			default: throw new NodeOperationError(ef.getNode(), `未知操作: ${operation}`);
		}
	}

	private static getResolvedSheetName(ef: IExecuteFunctions, itemIndex: number, paramName: 'sheetName' | 'worksheetName'): string {
		const inputMode = ef.getNodeParameter('inputMode', itemIndex) as string;
		let name = ef.getNodeParameter(paramName, itemIndex, '') as string;
		if (inputMode === 'filePath') {
			const optionParamName = paramName === 'sheetName' ? 'sheetNameOptions' : 'worksheetNameOptions';
			const nameOption = ef.getNodeParameter(optionParamName, itemIndex, '') as string;
			if (nameOption) name = nameOption;
		}
		return name;
	}

	private static async appendRow(ef: IExecuteFunctions, worksheet: ExcelJS.Worksheet, itemIndex: number) {
		const rowData = JSON.parse(ef.getNodeParameter('rowData', itemIndex) as string);
		const headers: string[] = [];
		worksheet.getRow(1).eachCell((cell, colNumber) => { headers[colNumber - 1] = cell.value as string; });
		worksheet.addRow(headers.map(h => rowData[h] || ''));
		return { success: true, operation: 'appendRow', rowNumber: worksheet.rowCount, data: rowData };
	}

	private static async readRows(ef: IExecuteFunctions, worksheet: ExcelJS.Worksheet, itemIndex: number) {
		const startRow = ef.getNodeParameter('startRow', itemIndex) as number;
		const endRow = ef.getNodeParameter('endRow', itemIndex, 0) as number;
		const headers: string[] = [];
		worksheet.getRow(1).eachCell((cell, colNumber) => { headers[colNumber - 1] = cell.value as string; });
		const rows: any[] = [];
		const lastRow = endRow > 0 ? endRow : worksheet.rowCount;
		for (let i = startRow; i <= lastRow; i++) {
			const rowData: any = {};
			worksheet.getRow(i).eachCell((cell, colNumber) => { if (headers[colNumber - 1]) rowData[headers[colNumber - 1]] = cell.value; });
			if (Object.keys(rowData).length) rows.push(rowData);
		}
		return rows;
	}

	private static async updateRow(ef: IExecuteFunctions, worksheet: ExcelJS.Worksheet, itemIndex: number) {
		const rowNumber = ef.getNodeParameter('rowNumber', itemIndex) as number;
		const updatedData = JSON.parse(ef.getNodeParameter('updatedData', itemIndex) as string);
		const row = worksheet.getRow(rowNumber);
		worksheet.getRow(1).eachCell((cell, colNumber) => {
			const h = cell.value as string;
		if (Object.prototype.hasOwnProperty.call(updatedData, h)) row.getCell(colNumber).value = updatedData[h];
		});
		row.commit();
		return { success: true, operation: 'updateRow', rowNumber };
	}

	private static async deleteRow(ef: IExecuteFunctions, worksheet: ExcelJS.Worksheet, itemIndex: number) {
		const rowNumber = ef.getNodeParameter('rowNumber', itemIndex) as number;
		worksheet.spliceRows(rowNumber, 1);
		return { success: true, operation: 'deleteRow', rowNumber };
	}

	private static async findRows(ef: IExecuteFunctions, worksheet: ExcelJS.Worksheet, itemIndex: number) {
		const inputMode = ef.getNodeParameter('inputMode', itemIndex) as string;
		const col = inputMode === 'filePath' 
			? ef.getNodeParameter('searchColumn', itemIndex) as string
			: ef.getNodeParameter('searchColumnManual', itemIndex) as string;
		const val = ef.getNodeParameter('searchValue', itemIndex) as string;
		const type = ef.getNodeParameter('matchType', itemIndex) as string;
		const returnRowNumbers = ef.getNodeParameter('returnRowNumbers', itemIndex, false) as boolean;
		let colIndex = -1;
		const headers: string[] = [];
		worksheet.getRow(1).eachCell((c, i) => {
			headers[i - 1] = c.value as string;
			if (c.value === col || String.fromCharCode(64 + i) === col) colIndex = i;
		});
		if (colIndex === -1) throw new NodeOperationError(ef.getNode(), `找不到欄位 ${col}`);
		const matches: any[] = [];
		const rowNumbers: number[] = [];
		worksheet.eachRow((row, rowNum) => {
			if (rowNum === 1) return;
			const cellVal = String(row.getCell(colIndex).value || '');
		const match = type === 'exact' ? cellVal === val : type === 'contains' ? cellVal.includes(val) : cellVal.startsWith(val);
			if (match) {
				rowNumbers.push(rowNum);
				if (!returnRowNumbers) {
					const data: any = { _rowNumber: rowNum };
					row.eachCell((c, i) => { if (headers[i - 1]) data[headers[i - 1]] = c.value; });
					matches.push(data);
				}
			}
		});
		
		if (returnRowNumbers) {
			return { rowNumbers };
		}
		return matches;
	}

	private static async insertRow(ef: IExecuteFunctions, worksheet: ExcelJS.Worksheet, itemIndex: number) {
		const rowNum = ef.getNodeParameter('rowNumber', itemIndex) as number;
		const data = JSON.parse(ef.getNodeParameter('rowData', itemIndex) as string);
		const headers: string[] = [];
		worksheet.getRow(1).eachCell((c, i) => { headers[i - 1] = c.value as string; });
		worksheet.insertRow(rowNum, headers.map(h => data[h] || ''));
		return { success: true, operation: 'insertRow', rowNumber: rowNum };
	}

	private static async createWorksheet(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number) {
		const name = ef.getNodeParameter('newSheetName', itemIndex) as string;
		if (workbook.getWorksheet(name)) throw new NodeOperationError(ef.getNode(), `工作表已存在`);
		const ws = workbook.addWorksheet(name);
		const data = JSON.parse(ef.getNodeParameter('initialData', itemIndex, '[]') as string);
		if (data.length) {
			const h = Object.keys(data[0]);
			ws.addRow(h);
			data.forEach((r: any) => ws.addRow(h.map(k => r[k] || '')));
		}
		return { success: true, operation: 'createWorksheet', name };
	}

	private static async deleteWorksheet(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number) {
		const name = ExcelCrud.getResolvedSheetName(ef, itemIndex, 'worksheetName');
		const ws = workbook.getWorksheet(name);
		if (!ws) throw new NodeOperationError(ef.getNode(), `找不到工作表`);
		workbook.removeWorksheet(ws.id);
		return { success: true, operation: 'deleteWorksheet', name };
	}

	private static async renameWorksheet(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number) {
		const oldN = ExcelCrud.getResolvedSheetName(ef, itemIndex, 'worksheetName');
		const newN = ef.getNodeParameter('newSheetName', itemIndex) as string;
		const ws = workbook.getWorksheet(oldN);
		if (!ws) throw new NodeOperationError(ef.getNode(), `找不到工作表`);
		ws.name = newN;
		return { success: true, operation: 'renameWorksheet', from: oldN, to: newN };
	}

	private static async listWorksheets(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number) {
		const hidden = ef.getNodeParameter('includeHidden', itemIndex, false) as boolean;
		const list: any[] = [];
		workbook.eachSheet((ws, id) => {
			if (hidden || (ws as any).state === 'visible' || !(ws as any).state) {
				list.push({ id, name: ws.name, rows: ws.rowCount });
			}
		});
		return { operation: 'listWorksheets', list };
	}

	private static async copyWorksheet(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number) {
		const srcN = ExcelCrud.getResolvedSheetName(ef, itemIndex, 'worksheetName');
		const newN = ef.getNodeParameter('newSheetName', itemIndex) as string;
		const src = workbook.getWorksheet(srcN);
		if (!src) throw new NodeOperationError(ef.getNode(), `找不到工作表`);
		const dst = workbook.addWorksheet(newN);
		src.eachRow((row, i) => {
			const r = dst.getRow(i);
			row.eachCell((c, j) => { r.getCell(j).value = c.value; r.getCell(j).style = { ...c.style }; });
			r.commit();
		});
		return { success: true, operation: 'copyWorksheet', from: srcN, to: newN };
	}

	private static async getWorksheetInfo(ef: IExecuteFunctions, workbook: ExcelJS.Workbook, itemIndex: number) {
		const name = ExcelCrud.getResolvedSheetName(ef, itemIndex, 'worksheetName');
		const ws = workbook.getWorksheet(name);
		if (!ws) throw new NodeOperationError(ef.getNode(), `找不到工作表`);
		const cols: any[] = [];
		ws.getRow(1).eachCell((c, i) => cols.push({ index: i, letter: String.fromCharCode(64 + i), header: c.value }));
		return { operation: 'getWorksheetInfo', name, rows: ws.rowCount, cols };
	}
}