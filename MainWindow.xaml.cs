using GenerateRaveSDSByAi.dataModel;
using GenerateRaveSDSByAi.utils;
using Microsoft.Win32;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Diagnostics;
using System.IO;
using System.Net.Http.Json;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Xceed.Words.NET;

namespace GenerateRaveSDSByAi
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string codelistInstructions;        
        string normaltypeInstructions;
        string addtypeInstructions;
        string fixtypeInstructions;
        string labtypeInstructions;
        AIconfig aiconfig;
        List<Xceed.Document.NET.Table> tables;

        List<Form> forms = new();                            //存储Form
        List<DataDictionary> dictList = new();   //存储反序列化的DataDictionary
        List<List<Field>> fieldList = new();                      //存储反序列化的Field
        string analytes;
        public MainWindow()
        {
            InitializeComponent();
            sp2.Visibility = Visibility.Collapsed;
            sp3.Visibility = Visibility.Collapsed;
            sp5.Visibility = Visibility.Collapsed;
            bd1.Visibility = Visibility.Collapsed;
            bd2.Visibility = Visibility.Collapsed;
            bd4.Visibility = Visibility.Collapsed;

            aiconfig = App.GlobalAIConfig;
            analytes = string.Join("\n", App.Analytes) + "\n##现在请根据下面的信息直接返回json，以[开始，以]结尾，用于后续的反序列化（无需输出额外信息，确保后续的过程顺利）。";

            codelistInstructions = File.ReadAllText(@"提示词\codelist解析.txt");            
            normaltypeInstructions = File.ReadAllText(@"提示词\普通表单的field解析.txt");
            addtypeInstructions = File.ReadAllText(@"提示词\ADD类型表单的field解析.txt");
            fixtypeInstructions = File.ReadAllText(@"提示词\FIX类型表单的field解析.txt");
            labtypeInstructions = File.ReadAllText(@"提示词\LAB类型表单的field解析.txt") + analytes ;
            

            tables = new();

            ExcelPackage.License.SetNonCommercialPersonal("ZhouHuawei");
        }

        private void btnUpload_Click(object sender, RoutedEventArgs e)
        {
            // 1. 弹出文件选择对话框
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word文件 (*.docx)|*.docx",
                Title = "选择需要上传的CRF文件",
                Multiselect = false
            };
            bool flag = false;

            // 2. 如果用户选择了文件，执行解析
            if (openFileDialog.ShowDialog() == true)
            {
                sp2.Visibility = Visibility.Collapsed;
                sp3.Visibility = Visibility.Collapsed;
                sp5.Visibility = Visibility.Collapsed;
                bd1.Visibility = Visibility.Collapsed;
                bd2.Visibility = Visibility.Collapsed; 
                bd4.Visibility = Visibility.Collapsed;

                try
                {
                    tbLog.Text = "正在检查表单......";
                    this.tables.Clear();

                    using (var document = DocX.Load(openFileDialog.FileName))
                    {
                        StringBuilder sb = new StringBuilder();
                        File.WriteAllText(@"运行日志\CRFchecklog.txt", "============开始检查格式");
                        foreach (var table in document.Tables)
                        {
                            if (table.Rows[0].Cells.Count == 1) continue;
                            sb.AppendLine(FormatCheckUtils.CheckCrfFormat(table));
                            this.tables.Add(table);
                            File.WriteAllText(@"运行日志\CRFchecklog.txt", sb.ToString());
                        }
                        tbLog.Text = sb.ToString();
                        flag = sb.ToString().IndexOf("错误") >= 0 ? false : true;
                    }
                }
                catch (Exception ex)
                {
                    tbLog.Text = $"解析失败：{ex.Message}";
                    this.tables.Clear();
                }
                finally
                {
                    if (flag)
                    {
                        sp2.Visibility = Visibility.Visible;
                        sp3.Visibility = Visibility.Visible; 
                        sp5.Visibility = Visibility.Visible;
                        bd1.Visibility = Visibility.Visible;
                        bd2.Visibility = Visibility.Visible;
                        bd4.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private void btnForm_Click(object sender, RoutedEventArgs e)
        {
            if (tables == null || tables.Count == 0)
            {
                tbLog.Text = "无表单数据可解析";
                return;
            }

            this.forms.Clear();
            btnUpload.IsEnabled = false;
            btnFieldList.IsEnabled = false;
            try
            {
                int i = 0;
                File.WriteAllText(@"运行日志\formLog.txt", "================解析时间：" + DateTime.Now.ToString() + "\r\n");
                foreach (var table in tables)
                {
                    Form currentForm = CRFAnalyzeUtils.GetForm(table);
                    this.forms.Add(currentForm);
                    i++;
                    File.AppendAllText(@"运行日志\formLog.txt", "成功提取表单：" + currentForm.Name + "\r\n");
                }
                this.tbLog.Text = $"成功提取{i}张表单！！";
            }
            catch (Exception ex)
            {
                tbLog.Text = $"解析失败：{ex.Message}";
                this.forms.Clear();
            }
            finally
            {
                btnUpload.IsEnabled = true;
                btnFieldList.IsEnabled = true;
            }

            tbLog.Text = "开始导出Form列表......";
            ExcelExporter.ExportFormToExcel(this.forms, @"运行日志\Forms.xlsx");
            tbLog.Inlines.Clear();
            Hyperlink link = new Hyperlink(new Run("成功导出Forms.xlsx"))
            {
                Foreground = Brushes.Blue,
                Cursor = Cursors.Hand,
                TextDecorations = TextDecorations.Underline
            };
            
            string excelFullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"运行日志\Forms.xlsx");
            link.Click += (s, e) =>
            {

                string folderPath = Path.GetDirectoryName(excelFullPath);
                if (Directory.Exists(folderPath))
                {
                    Process.Start(new ProcessStartInfo("explorer.exe")
                    {
                        Arguments = $"/select, \"{excelFullPath}\"",
                        UseShellExecute = true
                    });
                }
                else
                {
                    MessageBox.Show("❌ 文件夹【运行日志】不存在！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            };

            tbLog.Inlines.Add(link);

        }

        private async void btnCodeList_Click(object sender, RoutedEventArgs e)
        {
            if (tables == null || tables.Count == 0)
            {
                tbLog.Text = "无表单数据可解析";
                return;
            }
            this.dictList.Clear();
            List<string> layoutList = new List<string>();
            bool flag = false;
            string path = @"运行日志\codelistLog.txt";
            File.WriteAllText(path, $"====================开始解析字段的codelist！！！！" + DateTime.Now.ToString() + "\r\n");
            foreach (var table in tables)
            {
                try
                {
                    var datadicList = CRFAnalyzeUtils.GetDataDictionary(table);
                    tbLog.Text = string.Join("\n", datadicList);
                    layoutList.AddRange(datadicList);
                }
                catch (Exception ex)
                {
                    tbLog.Text = $"解析失败：{ex.Message}";
                    flag = true;
                }
            }
            File.AppendAllText(path, $"====================字段解析完毕！！！！" + "\r\n");
            tbLog.Text = "未从CRF中提取到有效的codeList，不会进行后续的AI解析......";

            if (flag) return;
            if (layoutList.Count == 0) return;
            tbLog.Text = "AI正在提取codeList......";


            //2. AI解析和反序列化
            int batchsize = aiconfig.batchSize;
            int round = (layoutList.Count + batchsize - 1) / batchsize;
            bool isAiProcessFailed = false;
            for (int i = 0; i < round; i++)
            {
                if (isAiProcessFailed)
                {
                    tbLog.Text += $"前批次AI解析失败，终止后续批次处理（剩余 {round - i} 批未处理）";
                    break; 
                }

                var curlist = layoutList.Skip(i * batchsize).Take(batchsize).ToList<string>();
                tbLog.Text = $"codelist一共被分为{round}批，AI正在处理第 {i + 1} 批 --- {curlist.Count} 条数据";
                await Task.Delay(2000);
                try
                {
                    string result = await CRFAnalyzeUtils.UsingAiTransferListToJson(codelistInstructions, curlist, aiconfig, path);

                    if (result.StartsWith("AI解析时发生异常"))
                    {
                        throw new Exception(result); // 包装异常，进入catch逻辑
                    }
                    tbLog.Text = "AI返回结果：" + result;
                    if (SafeJsonDeserializer.TryDeserializeFromAiText<DataDictionary>(result, out List<DataDictionary> dictlist, path))
                        this.dictList.AddRange(dictlist);

                }
                catch (Exception ex)
                {
                    string aiErrorMsg = $"第 {i + 1} 批AI解析失败：{ex.Message}";
                    tbLog.Text = aiErrorMsg;
                    File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {aiErrorMsg}\r\n异常详情：{ex.StackTrace}\r\n");
                    isAiProcessFailed = true;
                }

            }

            //3. 数据导出
            if (!isAiProcessFailed)
            {
                var list = this.dictList.GroupBy(g => g.OID).Select(g => g.First()).ToList();
                
                ExcelExporter.ExportDataDictionaryToExcel(list, @"运行日志\dataDictionary.xlsx");   
            }

        }

        private async void btnFieldList_Click(object sender, RoutedEventArgs e)
        {
            if (tables == null || tables.Count == 0)
            {
                tbLog.Text = "无表单数据可解析"; return;
            }
            this.fieldList.Clear();
            bool isAiProcessFailed = false;
            string path = @"运行日志\fieldlistLog.txt";
            File.WriteAllText(path, "====================" + DateTime.Now.ToString() + "\r\n");
            // ★ 新增：定义每批次拆分的大小，可根据AI接口性能调整（建议20-50，按需修改）
            int batchSize = aiconfig.batchSize;

            foreach (var table in tables)
            {
                if (isAiProcessFailed) break;
                if (table.Rows[0].Cells.Count < 2) continue;

                string formOIDandName = TextExtractor.GetCellValue(table, 0, 0);

                if (TextExtractor.ExtractNameAndOid(formOIDandName, out string formName, out string formOID))
                {
                    string crfType = table.Rows[0].Cells[1].Paragraphs[0].Text.Trim();
                    string result = "";
                    tbLog.Text = $"正在提取{formOIDandName}中的字段......";
                    File.AppendAllText(path, $"====================开始逐一解析{formOIDandName}中的字段！！！！" + DateTime.Now.ToString() + "\r\n");
                    try
                    {
                        // ★ 声明公共list变量，承接不同CRF类型的原始数据
                        List<string> list = new();
                        string currentInstruction = string.Empty;
                        // ★ 根据类型获取完整大List + 匹配对应的指令
                        switch (crfType)
                        {
                            case "ADD1":
                                list = CRFAnalyzeUtils.GetFieldInADD1Form(table, formOID);
                                currentInstruction = addtypeInstructions;
                                break;
                            case "ADD2":
                                list = CRFAnalyzeUtils.GetFieldInADD2Form(table, formOID);
                                currentInstruction = addtypeInstructions;
                                break;
                            case "FIX1":
                                list = CRFAnalyzeUtils.GetFieldInFIX1Form(table, formOID);
                                currentInstruction = fixtypeInstructions;
                                break;
                            case "FIX2":
                                list = CRFAnalyzeUtils.GetFieldInFIX2Form(table, formOID);
                                currentInstruction = fixtypeInstructions;
                                break;
                            case "LAB1":
                                list = CRFAnalyzeUtils.GetFieldInLAB1Form(table, formOID);
                                currentInstruction = labtypeInstructions;
                                break;
                            case "LAB2":
                                list = CRFAnalyzeUtils.GetFieldInLAB2Form(table, formOID);
                                currentInstruction = labtypeInstructions;
                                break;
                            default:
                                list = CRFAnalyzeUtils.GetFieldInNormalForm(table, formOID);
                                currentInstruction = normaltypeInstructions;
                                break;
                        }

                        // ★ ✅ 核心新增：循环拆分大List，分批调用AI
                        // 临时存储当前表单的所有分批解析结果
                        List<Field> currentFormAllFields = new();
                        // 计算总批次
                        int totalBatch = (int)Math.Ceiling((double)list.Count / batchSize);
                        File.AppendAllText(path, $"【拆分批次】总数据量：{list.Count}条，批次大小：{batchSize}条，总批次：{totalBatch}批\r\n");

                        for (int batchIndex = 0; batchIndex < totalBatch; batchIndex++)
                        {
                            if (isAiProcessFailed) break;                            
                            var batchList = list.Skip(batchIndex * batchSize).Take(batchSize).ToList();   
                            
                            result = await CRFAnalyzeUtils.UsingAiTransferListToJson(currentInstruction, batchList, aiconfig, path);
                            if (result.StartsWith("AI解析时发生异常"))
                            {
                                throw new Exception($"表单{formOIDandName}第{batchIndex + 1}批的字段解析失败：{result}");
                            }
                           
                            if (SafeJsonDeserializer.TryDeserializeFromAiText<Field>(result, out List<Field> batchFieldList, path))
                            {
                                for (int i = 0; i < batchFieldList.Count; i++)
                                {
                                    batchFieldList[i].UpdateField();
                                }
                                currentFormAllFields.AddRange(batchFieldList);
                                File.AppendAllText(path, $"第【{batchIndex + 1}/{totalBatch}】批解析成功，新增字段{batchFieldList.Count}个\r\n");
                            }
                        }

                        // 所有批次处理完成，将当前表单的完整结果加入全局集合
                        if (currentFormAllFields.Count > 0 && !isAiProcessFailed)
                        {
                            this.fieldList.Add(currentFormAllFields);                            
                        }

                    }
                    catch (Exception ex)
                    {
                        string aiErrorMsg = $"AI解析失败：{ex.Message}";
                        tbLog.Text = aiErrorMsg;
                        File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {aiErrorMsg}\r\n异常详情：{ex.StackTrace}\r\n");
                        isAiProcessFailed = true;

                        //及时保存已有的数据，防止丢失
                        if (this.fieldList.Count > 0)
                        {
                            ExcelExporter.ExportFieldToExcel(this.fieldList, @"运行日志\Field.xlsx");
                        }
                    }
                }
            }

            //3. 数据导出
            if (!isAiProcessFailed)
            {
                ExcelExporter.ExportFieldToExcel(this.fieldList, @"运行日志\Field.xlsx");
                await Task.Delay(5000);
                tbLog.Inlines.Clear();
                Hyperlink link = new Hyperlink(new Run("成功导出Field.xlsx"))
                {
                    Foreground = Brushes.Blue,
                    Cursor = Cursors.Hand,
                    TextDecorations = TextDecorations.Underline
                };
                string excelFullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"运行日志\Field.xlsx");
                link.Click += (s, e) =>
                {
                    string folderPath = Path.GetDirectoryName(excelFullPath);
                    if (Directory.Exists(folderPath))
                    {
                        Process.Start(new ProcessStartInfo("explorer.exe")
                        {
                            Arguments = $"/select, \"{excelFullPath}\"",
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        MessageBox.Show("❌ 文件夹【运行日志】不存在！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                };
                tbLog.Inlines.Add(link);
            }
        }

        //private async void btnFieldList_Click(object sender, RoutedEventArgs e)
        //{
        //    if (tables == null || tables.Count == 0)
        //    {
        //        tbLog.Text = "无表单数据可解析"; return;
        //    }
        //    this.fieldList.Clear();
        //    bool isAiProcessFailed = false;
        //    string path = @"运行日志\fieldlistLog.txt";
        //    File.WriteAllText(path, "====================" + DateTime.Now.ToString() + "\r\n");
        //    foreach (var table in tables)
        //    {                
        //        if (isAiProcessFailed) break;
        //        if (table.Rows[0].Cells.Count < 2) continue;

        //        string formOIDandName = TextExtractor.GetCellValue(table, 0, 0);

        //        if (TextExtractor.ExtractNameAndOid(formOIDandName, out string formName, out string formOID))
        //        {
        //            string crfType = table.Rows[0].Cells[1].Paragraphs[0].Text.Trim();
        //            string result = "";
        //            tbLog.Text = $"正在提取{formOIDandName}中的字段......";
        //            File.AppendAllText(path, $"====================开始逐一解析{formOIDandName}中的字段！！！！" + DateTime.Now.ToString() + "\r\n");
        //            try
        //            {
        //                switch (crfType)
        //                {
        //                    case "ADD1":
        //                        var list = CRFAnalyzeUtils.GetFieldInADD1Form(table, formOID);
        //                        //var list1 = list.Take(30).ToList();
        //                        //var list2 = list.Skip(30).Take(30).ToList();
        //                        //var list3 = list.Skip(60).ToList();
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(addtypeInstructions, list, aiconfig, path);

        //                        break;
        //                    case "ADD2":
        //                        list = CRFAnalyzeUtils.GetFieldInADD2Form(table, formOID);
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(addtypeInstructions, list, aiconfig, path);
        //                        break;
        //                    case "FIX1":
        //                        list = CRFAnalyzeUtils.GetFieldInFIX1Form(table, formOID);
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(fixtypeInstructions, list, aiconfig, path);
        //                        break;
        //                    case "FIX2":
        //                        list = CRFAnalyzeUtils.GetFieldInFIX2Form(table, formOID);
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(fixtypeInstructions, list, aiconfig, path);
        //                        break;
        //                    case "LAB1":
        //                        list = CRFAnalyzeUtils.GetFieldInLAB1Form(table, formOID);
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(labtypeInstructions, list, aiconfig, path);
        //                        break;
        //                    case "LAB2":
        //                        list = CRFAnalyzeUtils.GetFieldInLAB2Form(table, formOID);
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(labtypeInstructions, list, aiconfig, path);
        //                        break;
        //                    default:
        //                        list = CRFAnalyzeUtils.GetFieldInNormalForm(table, formOID);
        //                        result = await CRFAnalyzeUtils.UsingAiTransferListToJson(normaltypeInstructions, list, aiconfig, path);
        //                        break;
        //                }
        //                if (result.StartsWith("AI解析时发生异常"))
        //                {
        //                    throw new Exception(result);
        //                }

        //                tbLog.Text = "AI返回结果：" + result;

        //                if (SafeJsonDeserializer.TryDeserializeFromAiText<Field>(result, out List<Field> fieldlist, path))
        //                {
        //                    for (int i = 0; i < fieldlist.Count; i++)
        //                    {
        //                        fieldlist[i].UpdateField();
        //                    }
        //                    this.fieldList.Add(fieldlist);
        //                }

        //            }
        //            catch (Exception ex)
        //            {
        //                string aiErrorMsg = $"AI解析失败：{ex.Message}";
        //                tbLog.Text = aiErrorMsg;
        //                File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {aiErrorMsg}\r\n异常详情：{ex.StackTrace}\r\n");
        //                isAiProcessFailed = true;

        //                //及时保存已有的数据，防止丢失
        //                if (this.fieldList.Count > 0)
        //                {
        //                    ExcelExporter.ExportFieldToExcel(this.fieldList, @"运行日志\Field.xlsx");
        //                }
        //            }
        //        }
        //    }


        //    //3. 数据导出
        //    if (!isAiProcessFailed)
        //    {
        //        ExcelExporter.ExportFieldToExcel(this.fieldList, @"运行日志\Field.xlsx");                
        //        await Task.Delay(5000);
        //        tbLog.Inlines.Clear();
        //        Hyperlink link = new Hyperlink(new Run("成功导出Field.xlsx"))
        //        {
        //            Foreground = Brushes.Blue,          // 超链接蓝色
        //            Cursor = Cursors.Hand,              // 悬浮手型
        //            TextDecorations = TextDecorations.Underline // 下划线
        //        };
        //        //tbLog.Text = "成功导出Field.xlsx";
        //        string excelFullPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"运行日志\Field.xlsx");
        //        link.Click += (s, e) =>
        //        {

        //            string folderPath = Path.GetDirectoryName(excelFullPath);
        //            if (Directory.Exists(folderPath))
        //            {
        //                Process.Start(new ProcessStartInfo("explorer.exe")
        //                {
        //                    Arguments = $"/select, \"{excelFullPath}\"",
        //                    UseShellExecute = true
        //                });
        //            }
        //            else
        //            {
        //                MessageBox.Show("❌ 文件夹【运行日志】不存在！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
        //            }
        //        };

        //        tbLog.Inlines.Add(link);

        //    }


        //}


    }
}