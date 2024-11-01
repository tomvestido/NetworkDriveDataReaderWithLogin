using Scada.AddIn.Contracts;
using System;
using System.IO;
using System.Runtime.InteropServices;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Windows.Forms;

namespace ReadEnergyPricesNET4
{
    [AddInExtension("ReadEnergyPrices", "Odczyt cen energii z dysku sieciowego TPE")]
    public class ProjectServiceExtension : IProjectServiceExtension
    {
        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(ref NetResource netResource, string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);

        [StructLayout(LayoutKind.Sequential)]
        public struct NetResource
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpProvider;
        }

        public void Start(IProject context, IBehavior behavior)
        {
            // na potrzeby wew
            int actualYear = DateTime.Now.Year;
            int actualMonth = DateTime.Now.Month;
            int actualDay = DateTime.Now.Day;
      
            // nazwa pliku
            String date = $"{actualYear}{actualMonth:D2}{actualDay:D2}"; 
            String fileName = $"File_{date}.xlsx";

            // sciezka do pliku oraz dane do logowania
            string networkPath = @"yournetworkpath";
            string username = "user";
            string password = "password";
            string filePath = Path.Combine(networkPath, fileName);

            // wyzerowanie kodu bledu po uruchomieniu
            context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 0);

            var netResource = new NetResource
            {
                dwType = 1,
                lpRemoteName = networkPath
            };

            try
            {
                int result = WNetAddConnection2(ref netResource, password, username, 0);
                
                // weryfikacja czy logowanie odbylo sie poprawnie
                if (result != 0)
                {
                    context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 1);
                    return;
                }

                // weryfikacja czy plik istnieje
                if (!File.Exists(filePath))
                {
                    context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 2);
                    return;
                }

                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(file);
                    ISheet sheet = workbook.GetSheetAt(0);
                    IRow row = sheet.GetRow(19); 

                    // weryfikacja czy wiersz nie jest pusty
                    if (row == null)
                    {                       
                        context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 3);
                        return;
                    }

                    // pobranie godziny 24, aby byla w elemencie 0
                    ICell cellFirst = row.GetCell(26);
                    string cellValueStrFirst = cellFirst?.ToString() ?? "null";

                    // weryfikacja czy komorka nie jest pusta
                    if (cellFirst != null)
                    {
                        double cellValueFirst;
                        if (cellFirst.CellType == CellType.Numeric)
                        {
                            cellValueFirst = cellFirst.NumericCellValue;
                            context.VariableCollection[$"energyPricesToday[0]"].SetValue(0, cellValueFirst);
                        }
                        else if (cellFirst.CellType == CellType.String)
                        {
                            if (double.TryParse(cellValueStrFirst.Replace(",", "."), out cellValueFirst))
                            {
                                context.VariableCollection[$"energyPricesToday[0]"].SetValue(0, cellValueFirst);
                            }
                            else
                            {
                                context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 4);
                            }
                        }
                        else
                        {
                            context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 5);
                        }
                    }
                    else
                    {
                        context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 6);
                    }

                    // pobranie godzin 1-23
                    for (int i = 1; i < 24; i++)
                    {
                        ICell cell = row.GetCell(i + 2); 
                        string cellValueStr = cell?.ToString() ?? "null"; 
                        
                        // weryfikacja czy komorka nie jest pusta
                        if (cell != null)
                        {
                            double cellValue;
                            if (cell.CellType == CellType.Numeric)
                            {
                                cellValue = cell.NumericCellValue;
                                context.VariableCollection[$"energyPricesToday[{i}]"].SetValue(0, cellValue);
                            }
                            else if (cell.CellType == CellType.String)
                            {
                                if (double.TryParse(cellValueStr.Replace(",", "."), out cellValue))
                                {
                                    context.VariableCollection[$"energyPricesToday[{i}]"].SetValue(0, cellValue);
                                }
                                else
                                {                                   
                                    context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 4);
                                }
                            }
                            else
                            {
                                context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 5);
                            }
                        }
                        else
                        {
                            context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 6);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                context.VariableCollection["Harmonogram_arbitraz_kodBledu"].SetValue(0, 7);
            }
            finally
            {
                WNetCancelConnection2(networkPath, 0, true);
            }
        }

        public void Stop()
        {
            // Kod do wykonania przy zatrzymywaniu usługi
        }
    }
}
