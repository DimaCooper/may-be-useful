using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Excel = Microsoft.Office.Interop.Excel;

namespace CalcTab
{
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private Excel.Application excel = null;

		private Excel.Workbook wbParties = null;
		private Excel.Workbook wbNomenclatures = null;
		private Excel.Workbook wbMachineTools = null;
		private Excel.Workbook wbTimes = null;

		public MainWindow()
		{
			try
			{
				InitializeComponent();

				excel = new Excel.Application();

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (wbParties != null)
			{
				wbParties.Close();
				Marshal.ReleaseComObject(wbParties);
			}
			if (wbNomenclatures != null)
			{
				wbNomenclatures.Close();
				Marshal.ReleaseComObject(wbNomenclatures);
			}
			if (wbMachineTools != null)
			{
				wbMachineTools.Close();
				Marshal.ReleaseComObject(wbMachineTools);
			}
			if (wbTimes != null)
			{
				wbTimes.Close();
				Marshal.ReleaseComObject(wbTimes);
			}

			if (excel != null)
			{
				excel.Quit();
				Marshal.ReleaseComObject(excel);
			}
		}

		//Проверяем таблицы перед составлением расписания
		private void Save_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				#region Check Files
				if (wbParties == null)
				{
					throw new Exception("Select parties.xlsx");
				}
				if (wbNomenclatures == null)
				{
					throw new Exception("Select nomenclatures.xlsx");
				}
				if (wbMachineTools == null)
				{
					throw new Exception("Select machine_tools.xlsx");
				}
				if (wbTimes == null)
				{
					throw new Exception("Select times.xlsx");
				}

				if (wbParties.Sheets.Count <= 0)
				{
					throw new Exception("Wrong parties.xlsx");
				}
				if (wbNomenclatures.Sheets.Count <= 0)
				{
					throw new Exception("Wrong parties.xlsx");
				}
				if (wbMachineTools.Sheets.Count <= 0)
				{
					throw new Exception("Wrong parties.xlsx");
				}
				if (wbTimes.Sheets.Count <= 0)
				{
					throw new Exception("Wrong parties.xlsx");
				}
				#endregion Check Files

				List<Structures.Partie> parties = new List<Structures.Partie>();
				List<Structures.Nomenclature> nomenclatures = new List<Structures.Nomenclature>();
				List<Structures.MachineTool> machineTools = new List<Structures.MachineTool>();
				List<Structures.Time> times = new List<Structures.Time>();

				#region Read parties.xlsx
				// Считываем таблицу с партиями сырья
				try
				{
					Excel._Worksheet xlWorksheet = wbParties.Sheets[1];
					Excel.Range xlRange = xlWorksheet.UsedRange;
					object[,] xlValues = (object[,])xlRange.Value2;

					int rows = xlRange.Rows.Count;
					int columns = xlRange.Columns.Count;

					for (int i = 0; i < rows; ++i)
					{
						Structures.Partie partie = null;

						for (int j = 0; j < columns; ++j)
						{
							if (xlValues[i + 1, j + 1] is double cell)
							{
								if (partie == null)
								{
									partie = new Structures.Partie();
								}
								if (j == 0)
								{
									partie.id = cell;
								}
								else if (j == 1)
								{
									partie.nomenclatureID = cell;
								}
							}
						}

						if (partie != null)
						{
							parties.Add(partie);
						}
					}

					Marshal.ReleaseComObject(xlRange);
					Marshal.ReleaseComObject(xlWorksheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Wrong parties.xlsx\n\n{ex.Message}");
				}
				#endregion Read parties.xlsx

				#region Read nomenclatures.xlsx
				// Считываем таблицу с наименованием сырья
				try
				{
					Excel._Worksheet xlWorksheet = wbNomenclatures.Sheets[1];
					Excel.Range xlRange = xlWorksheet.UsedRange;
					object[,] xlValues = (object[,])xlRange.Value2;

					int rows = xlRange.Rows.Count;
					int columns = xlRange.Columns.Count;

					for (int i = 0; i < rows; ++i)
					{
						Structures.Nomenclature nomenclature = null;

						for (int j = 0; j < columns; ++j)
						{
							if (xlValues[i + 1, j + 1] is double cell)
							{
								if (nomenclature == null)
								{
									nomenclature = new Structures.Nomenclature();
								}
								if (j == 0)
								{
									nomenclature.id = cell;
								}
							}
							if (xlValues[i + 1, j + 1] is string cellNomenclature)
							{
								if (nomenclature != null && j == 1)
								{
									nomenclature.nomenclature = cellNomenclature;
								}
							}
						}

						if (nomenclature != null)
						{
							nomenclatures.Add(nomenclature);
						}
					}

					Marshal.ReleaseComObject(xlRange);
					Marshal.ReleaseComObject(xlWorksheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Wrong nomenclatures.xlsx\n\n{ex.Message}");
				}
				#endregion Read nomenclatures.xlsx

				#region Read machine_tools.xlsx
				// Считываем таблицу с оборудование
				try
				{
					Excel._Worksheet xlWorksheet = wbMachineTools.Sheets[1];
					Excel.Range xlRange = xlWorksheet.UsedRange;
					object[,] xlValues = (object[,])xlRange.Value2;

					int rows = xlRange.Rows.Count;
					int columns = xlRange.Columns.Count;

					for (int i = 0; i < rows; ++i)
					{
						Structures.MachineTool machineTool = null;

						for (int j = 0; j < columns; ++j)
						{
							if (xlValues[i + 1, j + 1] is double cell)
							{
								if (machineTool == null)
								{
									machineTool = new Structures.MachineTool();
								}
								if (j == 0)
								{
									machineTool.id = cell;
								}
							}
							if (xlValues[i + 1, j + 1] is string cellName)
							{
								if (machineTool != null && j == 1)
								{
									machineTool.name = cellName;
								}
							}
						}

						if (machineTool != null)
						{
							machineTools.Add(machineTool);
						}
					}

					Marshal.ReleaseComObject(xlRange);
					Marshal.ReleaseComObject(xlWorksheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Wrong nomenclatures.xlsx\n\n{ex.Message}");
				}
				#endregion Read nomenclatures.xlsx

				#region Read times.xlsx
				// Считываем таблицу с таймингами
				try
				{
					Excel._Worksheet xlWorksheet = wbTimes.Sheets[1];
					Excel.Range xlRange = xlWorksheet.UsedRange;
					object[,] xlValues = (object[,])xlRange.Value2;

					int rows = xlRange.Rows.Count;
					int columns = xlRange.Columns.Count;

					for (int i = 0; i < rows; ++i)
					{
						Structures.Time time = null;

						for (int j = 0; j < columns; ++j)
						{
							if (xlValues[i + 1, j + 1] is double cell)
							{
								if (time == null)
								{
									time = new Structures.Time();
								}
								if (j == 0)
								{
									time.machineToolID = cell;
								}
								else if (j == 1)
								{
									time.nomenclatureID = cell;
								}
								else if (j == 2)
								{
									time.operationTime = cell;
								}
							}
						}

						if (time != null)
						{
							times.Add(time);
						}
					}

					Marshal.ReleaseComObject(xlRange);
					Marshal.ReleaseComObject(xlWorksheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Wrong times.xlsx\n\n{ex.Message}");
				}
				#endregion Read times.xlsx

				List<Structures.Result> results = new List<Structures.Result>();

				#region Calculate Result

				// Расчитываем время обработки сырья
				foreach (Structures.Partie partie in parties)
				{
					List<Structures.MachineTool> machineToolsSupport = new List<Structures.MachineTool>();

					foreach (Structures.Time time in times)
					{
						if (partie.nomenclatureID == time.nomenclatureID)
						{
							foreach (Structures.MachineTool machineTool in machineTools)
							{
								if (time.machineToolID == machineTool.id)
								{
									machineToolsSupport.Add(new Structures.MachineTool
									{
										id = machineTool.id,
										name = machineTool.name,
										work = machineTool.work,
										operationTime = time.operationTime
									});
								}
							}
						}
					}

					Structures.MachineTool machineToolUse = machineToolsSupport[0];

					foreach (Structures.MachineTool machineToolSupport in machineToolsSupport)
					{
						if (machineToolSupport.work < machineToolUse.work)
						{
							machineToolUse = machineToolSupport;
						}
					}

					string partieName = string.Empty;
					foreach (Structures.Nomenclature nomenclature in nomenclatures)
					{
						if (partie.nomenclatureID == nomenclature.id)
						{
							partieName = nomenclature.nomenclature;
							break;
						}
					}

					results.Add(new Structures.Result
					{
						partieID = partie.id,
						partieName = partieName,
						machineToolName = machineToolUse.name,
						begin = machineToolUse.work,
						end = machineToolUse.work += machineToolUse.operationTime
					});

					foreach (Structures.MachineTool machineTool in machineTools)
					{
						if (machineToolUse.id == machineTool.id)
						{
							machineTool.work = machineToolUse.work;
							break;
						}
					}
				}
				#endregion Calculate Result

				// Добавляем расчет в таблицу и сохраняем документ
				Excel.Workbook wbResult = excel.Workbooks.Add(Type.Missing);
				Excel.Worksheet wsResult = wbResult.ActiveSheet;

				for (int i = 0; i < results.Count; ++i)
				{
					Structures.Result result = results[i];

					Excel.Range xlRange = null;

					xlRange = wsResult.get_Range($"A{i + 1}", Type.Missing);
					xlRange.Value2 = result.partieID;

					xlRange = wsResult.get_Range($"B{i + 1}", Type.Missing);
					xlRange.Value2 = result.partieName;

					xlRange = wsResult.get_Range($"C{i + 1}", Type.Missing);
					xlRange.Value2 = result.machineToolName;

					xlRange = wsResult.get_Range($"D{i + 1}", Type.Missing);
					xlRange.Value2 = result.begin;

					xlRange = wsResult.get_Range($"E{i + 1}", Type.Missing);
					xlRange.Value2 = result.end;

					Marshal.ReleaseComObject(xlRange);
				}

				SaveFileDialog saveDialog = new SaveFileDialog
				{
					OverwritePrompt = false,
					FileName = "result.xlsx"
				};
				bool? dialogResult = saveDialog.ShowDialog();
				if (dialogResult.HasValue && dialogResult.Value)
				{
					wbResult.SaveAs(saveDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
				}

				wbResult.Close();

				Marshal.ReleaseComObject(wsResult);
				Marshal.ReleaseComObject(wbResult);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		//Обрабатываем клик по кнопке выбора таблицы партии
		private void Parties_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog ofd = new OpenFileDialog();
				bool? result = ofd.ShowDialog();
				if (result.HasValue && result.Value)
				{
					wbParties = excel.Workbooks.Open(ofd.FileName);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		//Обрабатываем клик по кнопке выбора таблицы сырья
		private void Nomenclatures_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog ofd = new OpenFileDialog();
				bool? result = ofd.ShowDialog();
				if (result.HasValue && result.Value)
				{
					wbNomenclatures = excel.Workbooks.Open(ofd.FileName);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		//Обрабатываем клик по кнопке выбора таблицы оборудования
		private void MachineTools_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog ofd = new OpenFileDialog();
				bool? result = ofd.ShowDialog();
				if (result.HasValue && result.Value)
				{
					wbMachineTools = excel.Workbooks.Open(ofd.FileName);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		//Обрабатываем клик по кнопке выбора таблицы с таймингами
		private void Times_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				OpenFileDialog ofd = new OpenFileDialog();
				bool? result = ofd.ShowDialog();
				if (result.HasValue && result.Value)
				{
					wbTimes = excel.Workbooks.Open(ofd.FileName);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}
	}
}
