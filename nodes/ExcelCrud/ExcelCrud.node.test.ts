import { IExecuteFunctions, ILoadOptionsFunctions, INodeExecutionData, NodeOperationError } from 'n8n-workflow';
import { ExcelCrud } from './ExcelCrud.node';
import * as ExcelJS from 'exceljs';
import * as fs from 'fs/promises';

// Mock modules
jest.mock('fs/promises');
jest.mock('exceljs');

describe('ExcelCrud Node', () => {
	let excelCrud: ExcelCrud;
	let mockExecuteFunctions: jest.Mocked<IExecuteFunctions>;
	let mockLoadOptionsFunctions: jest.Mocked<ILoadOptionsFunctions>;
	let mockWorkbook: jest.Mocked<ExcelJS.Workbook>;
	let mockWorksheet: jest.Mocked<ExcelJS.Worksheet>;

	beforeEach(() => {
		excelCrud = new ExcelCrud();

		// 建立 mock worksheet
		mockWorksheet = {
			name: 'Sheet1',
			getRow: jest.fn(),
			addRow: jest.fn(),
			eachRow: jest.fn(),
			spliceRows: jest.fn(),
			insertRow: jest.fn(),
			getCell: jest.fn(),
		} as any;

		// 設置 readonly 屬性
		Object.defineProperty(mockWorksheet, 'rowCount', {
			get: jest.fn().mockReturnValue(10),
			configurable: true,
		});
		Object.defineProperty(mockWorksheet, 'columnCount', {
			get: jest.fn().mockReturnValue(5),
			configurable: true,
		});
		Object.defineProperty(mockWorksheet, 'id', {
			get: jest.fn().mockReturnValue(1),
			configurable: true,
		});

		// 建立 mock workbook
		mockWorkbook = {
			worksheets: [mockWorksheet],
			getWorksheet: jest.fn().mockReturnValue(mockWorksheet),
			addWorksheet: jest.fn().mockReturnValue(mockWorksheet),
			removeWorksheet: jest.fn(),
			eachSheet: jest.fn(),
			xlsx: {
				readFile: jest.fn().mockResolvedValue(undefined),
				load: jest.fn().mockResolvedValue(undefined),
				writeFile: jest.fn().mockResolvedValue(undefined),
				writeBuffer: jest.fn().mockResolvedValue(Buffer.from('mock-excel-data')),
			},
		} as any;

		// Mock ExcelJS.Workbook constructor
		(ExcelJS.Workbook as jest.MockedClass<typeof ExcelJS.Workbook>).mockImplementation(() => mockWorkbook);

		// 建立 mock execute functions
		mockExecuteFunctions = {
			getInputData: jest.fn().mockReturnValue([{ json: {} }]),
			getNodeParameter: jest.fn(),
			getNode: jest.fn().mockReturnValue({ name: 'ExcelCrud', type: 'n8n-nodes-excel-crud.excelCrud' }),
			helpers: {
				assertBinaryData: jest.fn(),
				getBinaryDataBuffer: jest.fn().mockResolvedValue(Buffer.from('mock-data')),
				prepareBinaryData: jest.fn().mockResolvedValue({ data: 'mock-binary', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', fileName: 'test.xlsx' }),
			},
			continueOnFail: jest.fn().mockReturnValue(false),
		} as any;

		// 建立 mock load options functions
		mockLoadOptionsFunctions = {
			getNodeParameter: jest.fn(),
		} as any;

		// Mock fs.access
		(fs.access as jest.Mock).mockResolvedValue(undefined);
	});

	afterEach(() => {
		jest.clearAllMocks();
	});

	describe('Node Properties', () => {
		it('should have correct node description', () => {
			expect(excelCrud.description.displayName).toBe('Excel CRUD');
			expect(excelCrud.description.name).toBe('excelCrud');
			expect(excelCrud.description.group).toContain('transform');
			expect(excelCrud.description.version).toBe(1);
		});

		it('should have all required properties', () => {
			const properties = excelCrud.description.properties;
			expect(properties).toBeDefined();
			expect(properties.length).toBeGreaterThan(0);

			const propertyNames = properties.map(p => p.name);
			expect(propertyNames).toContain('resource');
			expect(propertyNames).toContain('inputMode');
			expect(propertyNames).toContain('operation');
		});

		it('should have row operations', () => {
			const operationProperty = excelCrud.description.properties.find(p => p.name === 'operation');
			expect(operationProperty).toBeDefined();
			if (operationProperty && 'options' in operationProperty) {
				const operations = (operationProperty.options as any[]).map((o: any) => o.value);
				expect(operations).toContain('appendRow');
				expect(operations).toContain('readRows');
				expect(operations).toContain('updateRow');
				expect(operations).toContain('deleteRow');
				expect(operations).toContain('findRows');
				expect(operations).toContain('insertRow');
			}
		});
	});

	describe('Load Options Methods', () => {
		describe('getWorksheets', () => {
			it('should return empty array when inputMode is not filePath', async () => {
				mockLoadOptionsFunctions.getNodeParameter = jest.fn()
					.mockReturnValueOnce('binaryData'); // inputMode

				const result = await excelCrud.methods.loadOptions.getWorksheets.call(mockLoadOptionsFunctions);
				expect(result).toEqual([]);
			});

			it('should return empty array when filePath is not provided', async () => {
				mockLoadOptionsFunctions.getNodeParameter = jest.fn()
					.mockReturnValueOnce('filePath') // inputMode
					.mockReturnValueOnce(''); // filePath

				const result = await excelCrud.methods.loadOptions.getWorksheets.call(mockLoadOptionsFunctions);
				expect(result).toEqual([]);
			});

			it('should return worksheets when file exists', async () => {
				mockLoadOptionsFunctions.getNodeParameter = jest.fn()
					.mockReturnValueOnce('filePath') // inputMode
					.mockReturnValueOnce('/test/file.xlsx'); // filePath

				mockWorkbook.worksheets = [
					{ name: 'Sheet1' } as any,
					{ name: 'Sheet2' } as any,
				];

				const result = await excelCrud.methods.loadOptions.getWorksheets.call(mockLoadOptionsFunctions);

				expect(result).toEqual([
					{ name: 'Sheet1', value: 'Sheet1' },
					{ name: 'Sheet2', value: 'Sheet2' },
				]);
			});

			it('should return error message when file cannot be read', async () => {
				mockLoadOptionsFunctions.getNodeParameter = jest.fn()
					.mockReturnValueOnce('filePath')
					.mockReturnValueOnce('/test/nonexistent.xlsx');

				(fs.access as jest.Mock).mockRejectedValueOnce(new Error('File not found'));

				const result = await excelCrud.methods.loadOptions.getWorksheets.call(mockLoadOptionsFunctions);

				expect(result[0].name).toContain('Error');
			});
		});

		describe('getColumns', () => {
			it('should return columns for filePath mode', async () => {
				mockLoadOptionsFunctions.getNodeParameter = jest.fn()
					.mockReturnValueOnce('filePath') // inputMode
					.mockReturnValueOnce('/test/file.xlsx') // filePath
					.mockReturnValueOnce('Sheet1'); // sheetName

				const mockRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
						callback({ value: 'Age' }, 2);
						callback({ value: 'Email' }, 3);
					}),
				};

				mockWorksheet.getRow = jest.fn().mockReturnValue(mockRow);

				const result = await excelCrud.methods.loadOptions.getColumns.call(mockLoadOptionsFunctions);

				expect(result).toEqual([
					{ name: 'Name', value: 'Name' },
					{ name: 'Age', value: 'Age' },
					{ name: 'Email', value: 'Email' },
				]);
			});

			it('should handle binaryData mode appropriately', async () => {
				mockLoadOptionsFunctions.getNodeParameter = jest.fn()
					.mockReturnValueOnce('binaryData');

				const result = await excelCrud.methods.loadOptions.getColumns.call(mockLoadOptionsFunctions);

				expect(result.length).toBeGreaterThan(0);
				expect(result[0].name).toContain('file path mode');
			});
		});
	});

	describe('Row Operations - File Path Mode', () => {
		beforeEach(() => {
			mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					inputMode: 'filePath',
					filePath: '/test/file.xlsx',
					sheetNameOptions: 'Sheet1',
					autoSave: true,
				};
				return params[name] ?? defaultValue;
			}) as any;
		});

		describe('Append Row', () => {
			it('should append row successfully', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'appendRow',
						rowData: JSON.stringify({ Name: 'John', Age: 30 }),
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
						callback({ value: 'Age' }, 2);
					}),
				};

				mockWorksheet.getRow = jest.fn().mockReturnValue(mockHeaderRow);
			Object.defineProperty(mockWorksheet, 'rowCount', {
				get: jest.fn().mockReturnValue(5),
				configurable: true,
			});
				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('success', true);
				expect(result[0][0].json).toHaveProperty('operation', 'appendRow');
				expect(mockWorksheet.addRow).toHaveBeenCalled();
			});
		});

		describe('Read Rows', () => {
			it('should read rows successfully', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'readRows',
						startRow: 2,
						endRow: 0,
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
						callback({ value: 'Age' }, 2);
					}),
				};

				const mockDataRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'John' }, 1);
						callback({ value: 30 }, 2);
					}),
				};

mockWorksheet.getRow = jest.fn((rowNum: number) => {
				return rowNum === 1 ? mockHeaderRow : mockDataRow;
			}) as any;
			Object.defineProperty(mockWorksheet, 'rowCount', {
				get: jest.fn().mockReturnValue(3),
				configurable: true,
			});

				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0].length).toBeGreaterThan(0);
				expect(result[0][0].json).toHaveProperty('Name');
				expect(result[0][0].json).toHaveProperty('Age');
			});
		});

		describe('Update Row', () => {
			it('should update row successfully', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'updateRow',
						rowNumber: 2,
						updatedData: JSON.stringify({ Age: 31 }),
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockCell = { value: null };
				const mockRow = {
					getCell: jest.fn().mockReturnValue(mockCell),
					commit: jest.fn(),
					eachCell: jest.fn(),
				};

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
						callback({ value: 'Age' }, 2);
					}),
				};

mockWorksheet.getRow = jest.fn((rowNum: number) => {
				return rowNum === 1 ? mockHeaderRow : mockRow;
			}) as any;

				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('success', true);
				expect(result[0][0].json).toHaveProperty('operation', 'updateRow');
				expect(mockRow.commit).toHaveBeenCalled();
			});
		});

		describe('Delete Row', () => {
			it('should delete row successfully', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'deleteRow',
						rowNumber: 2,
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('success', true);
				expect(result[0][0].json).toHaveProperty('operation', 'deleteRow');
				expect(mockWorksheet.spliceRows).toHaveBeenCalledWith(2, 1);
			});
		});

		describe('Find Rows', () => {
			it('should find rows with exact match and return data', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'findRows',
						searchColumn: 'Name',
						searchValue: 'John',
						matchType: 'exact',
						returnRowNumbers: false,
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
						callback({ value: 'Age' }, 2);
					}),
				};

				mockWorksheet.getRow = jest.fn().mockReturnValue(mockHeaderRow);

			mockWorksheet.eachRow = jest.fn((callback: (row: any, rowNum: number) => void) => {
				const mockRow1 = {
					getCell: jest.fn().mockReturnValue({ value: 'Name' }),
					eachCell: jest.fn((cb: any) => {
						cb({ value: 'Name' }, 1);
						cb({ value: 'Age' }, 2);
					}),
				};
				const mockRow2 = {
					getCell: jest.fn().mockReturnValue({ value: 'John' }),
					eachCell: jest.fn((cb: any) => {
						cb({ value: 'John' }, 1);
						cb({ value: 30 }, 2);
					}),
				};
				callback(mockRow1, 1);
				callback(mockRow2, 2);
			}) as any;

			const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0].length).toBeGreaterThan(0);
				expect(result[0][0].json).toHaveProperty('_rowNumber');
			});

			it('should return row numbers when returnRowNumbers is true', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'findRows',
						searchColumn: 'Status',
						searchValue: 'Active',
						matchType: 'exact',
						returnRowNumbers: true,
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Status' }, 1);
					}),
				};

				mockWorksheet.getRow = jest.fn().mockReturnValue(mockHeaderRow);

			mockWorksheet.eachRow = jest.fn((callback: (row: any, rowNum: number) => void) => {
				callback({ getCell: jest.fn().mockReturnValue({ value: 'Status' }) }, 1);
				callback({ getCell: jest.fn().mockReturnValue({ value: 'Active' }) }, 2);
				callback({ getCell: jest.fn().mockReturnValue({ value: 'Inactive' }) }, 3);
				callback({ getCell: jest.fn().mockReturnValue({ value: 'Active' }) }, 5);
			}) as any;

			const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0][0].json.rowNumbers).toContain(2);
				expect(result[0][0].json.rowNumbers).toContain(5);
			});

			it('should find rows with contains match', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'findRows',
						searchColumn: 'Name',
						searchValue: 'Jo',
						matchType: 'contains',
						returnRowNumbers: false,
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
					}),
				};

				mockWorksheet.getRow = jest.fn().mockReturnValue(mockHeaderRow);

			mockWorksheet.eachRow = jest.fn((callback: (row: any, rowNum: number) => void) => {
				callback({ getCell: jest.fn().mockReturnValue({ value: 'Name' }) }, 1);
				callback({
					getCell: jest.fn().mockReturnValue({ value: 'John' }),
					eachCell: jest.fn((cb: any) => cb({ value: 'John' }, 1))
				}, 2);
			}) as any;
			const result = await excelCrud.execute.call(mockExecuteFunctions);
				expect(result[0].length).toBeGreaterThan(0);
			});
		});

		describe('Insert Row', () => {
			it('should insert row successfully', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'row',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						sheetNameOptions: 'Sheet1',
						operation: 'insertRow',
						rowNumber: 2,
						rowData: JSON.stringify({ Name: 'Jane', Age: 25 }),
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				const mockHeaderRow = {
					eachCell: jest.fn((callback) => {
						callback({ value: 'Name' }, 1);
						callback({ value: 'Age' }, 2);
					}),
				};

				mockWorksheet.getRow = jest.fn().mockReturnValue(mockHeaderRow);

				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('success', true);
				expect(result[0][0].json).toHaveProperty('operation', 'insertRow');
				expect(mockWorksheet.insertRow).toHaveBeenCalledWith(2, expect.any(Array));
			});
		});
	});

	describe('Worksheet Operations', () => {
		describe('Create Worksheet', () => {
			it('should create worksheet without initial data', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'worksheet',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						worksheetOperation: 'createWorksheet',
						newSheetName: 'NewSheet',
						initialData: '[]',
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				mockWorkbook.getWorksheet = jest.fn().mockReturnValue(null);

				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('success', true);
				expect(result[0][0].json).toHaveProperty('operation', 'createWorksheet');
				expect(mockWorkbook.addWorksheet).toHaveBeenCalledWith('NewSheet');
			});

			it('should create worksheet with initial data', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'worksheet',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						worksheetOperation: 'createWorksheet',
						newSheetName: 'DataSheet',
						initialData: JSON.stringify([
							{ Name: 'John', Age: 30 },
							{ Name: 'Jane', Age: 25 },
						]),
						autoSave: true,
					};
					return params[name] ?? defaultValue;
				});

				mockWorkbook.getWorksheet = jest.fn().mockReturnValue(null);

				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(mockWorksheet.addRow).toHaveBeenCalled();
			});
		});

		describe('Delete Worksheet', () => {
			it('should delete worksheet successfully', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'worksheet',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						worksheetOperation: 'deleteWorksheet',
						worksheetNameOptions: 'Sheet1',
						autoSave: true,
					};
					return params[name] ?? defaultValue;
			}) as any;

			Object.defineProperty(mockWorksheet, 'id', {
				get: jest.fn().mockReturnValue(1),
				configurable: true,
			});
				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('success', true);
				expect(mockWorkbook.removeWorksheet).toHaveBeenCalledWith(1);
			});
		});

		describe('List Worksheets', () => {
			it('should list all visible worksheets', async () => {
				mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
					const params: Record<string, any> = {
						resource: 'worksheet',
						inputMode: 'filePath',
						filePath: '/test/file.xlsx',
						worksheetOperation: 'listWorksheets',
						includeHidden: false,
						autoSave: true,
					};
					return params[name] ?? defaultValue;
			}) as any;

			mockWorkbook.eachSheet = jest.fn((callback: (sheet: any, id: number) => void) => {
				callback({ name: 'Sheet1', rowCount: 10, state: 'visible' } as any, 1);
				callback({ name: 'Sheet2', rowCount: 20, state: 'visible' } as any, 2);
			}) as any;
				const result = await excelCrud.execute.call(mockExecuteFunctions);

				expect(result[0]).toHaveLength(1);
				expect(result[0][0].json).toHaveProperty('operation', 'listWorksheets');
				expect(result[0][0].json).toHaveProperty('list');
			});
		});
	});

	describe('Error Handling', () => {
		it('should handle missing worksheet error', async () => {
			mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					inputMode: 'filePath',
					filePath: '/test/file.xlsx',
					sheetNameOptions: 'NonExistent',
					operation: 'readRows',
					startRow: 2,
					endRow: 0,
				};
				return params[name] ?? defaultValue;
			});

			mockWorkbook.getWorksheet = jest.fn().mockReturnValue(null);

			await expect(excelCrud.execute.call(mockExecuteFunctions)).rejects.toThrow();
		});

		it('should handle file read error', async () => {
			mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					inputMode: 'filePath',
					filePath: '/nonexistent/file.xlsx',
					operation: 'readRows',
					sheetNameOptions: 'Sheet1',
					startRow: 2,
					endRow: 0,
				};
				return params[name] ?? defaultValue;
			});

			(fs.access as jest.Mock).mockRejectedValueOnce(new Error('ENOENT'));

			await expect(excelCrud.execute.call(mockExecuteFunctions)).rejects.toThrow();
		});

		it('should continue on fail when configured', async () => {
			mockExecuteFunctions.continueOnFail = jest.fn().mockReturnValue(true);
			mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					inputMode: 'filePath',
					filePath: '/test/file.xlsx',
					sheetNameOptions: 'NonExistent',
					operation: 'readRows',
					startRow: 2,
					endRow: 0,
				};
				return params[name] ?? defaultValue;
			});

			mockWorkbook.getWorksheet = jest.fn().mockReturnValue(null);

			const result = await excelCrud.execute.call(mockExecuteFunctions);

			expect(result[0]).toHaveLength(1);
			expect(result[0][0].json).toHaveProperty('error');
		});
	});

	describe('Binary Data Mode', () => {
		it('should handle binary data input', async () => {
			mockExecuteFunctions.getNodeParameter = jest.fn((name: string, itemIndex: number, defaultValue?: any) => {
				const params: Record<string, any> = {
					resource: 'row',
					inputMode: 'binaryData',
					binaryPropertyName: 'data',
					sheetName: 'Sheet1',
					operation: 'readRows',
					startRow: 2,
					endRow: 0,
				};
				return params[name] ?? defaultValue;
			});

			const mockHeaderRow = {
				eachCell: jest.fn((callback) => {
					callback({ value: 'Name' }, 1);
				}),
			};

			mockWorksheet.getRow = jest.fn().mockReturnValue(mockHeaderRow);
			Object.defineProperty(mockWorksheet, 'rowCount', {
				get: jest.fn().mockReturnValue(2),
				configurable: true,
			});

			const result = await excelCrud.execute.call(mockExecuteFunctions);

			expect(result[0].length).toBeGreaterThan(0);
			expect(mockExecuteFunctions.helpers.assertBinaryData).toHaveBeenCalled();
			expect(mockExecuteFunctions.helpers.getBinaryDataBuffer).toHaveBeenCalled();
		});
	});
});
