using Sunny.UI;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Drawing;
using static NPOI.HSSF.Util.HSSFColor;
using System.Data.SQLite;
using System.Data;
using static Org.BouncyCastle.Crypto.Digests.SkeinEngine;

namespace SumRDTools
{
    public partial class Form1 : UIForm
    {
        //定义Excel处理和日志打印的委托
        public delegate void DealExcelAndPrintLogDelegate(DirectoryInfo directoryInfo);
        DealExcelAndPrintLogDelegate dealExcelAndPrintLogDelegate;
        public Form1()
        {
            InitializeComponent();
            //初始化Excel处理及日志打印委托
            dealExcelAndPrintLogDelegate = new DealExcelAndPrintLogDelegate(dealExcelAndPrintLogFun);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //如果配置文件中有路径，直接读取路径赋值到路径文件框中
            String defaultOpenPath = ConfigFileUtil.getConfigParam("DEFALUT_PARAM", "FILE_PATH");
            if (!String.IsNullOrEmpty(defaultOpenPath))
            {
                this.folderPathText.Text = defaultOpenPath;
            }
        }

        //选择文件夹目录
        private void chooseFolderBtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            String defaultOpenPath = ConfigFileUtil.getConfigParam("DEFALUT_PARAM", "FILE_PATH");
            if (!String.IsNullOrEmpty(defaultOpenPath)) {
                folderBrowserDialog.SelectedPath = defaultOpenPath;
            }
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK) { 
                this.folderPathText.Text = folderBrowserDialog.SelectedPath;
                //如果路径不一致的时候，更新到配置文件
                if (folderBrowserDialog.SelectedPath != defaultOpenPath) {
                    ConfigFileUtil.SetValue("DEFALUT_PARAM", "FILE_PATH", folderBrowserDialog.SelectedPath);
                }
            }
        }

        ////获取指定目录下的Excel，然后汇总对应的数据
        private void summaryBtn_Click(object sender, EventArgs e)
        {
            //清空日志，以便打印新的
            this.logTextBox.Clear();
            this.errorLogRichTextBox.Clear();

            String filesPath = this.folderPathText.Text;
            if (String.IsNullOrEmpty(filesPath))
            {
                MessageBox.Show("请选择需要汇总数据的目录！");
            }
            else
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(filesPath);
                if (directoryInfo.Exists == false)
                {
                    MessageBox.Show("请重新选择目录，该目录无效！");
                }
                else
                {
                    //调用委托开始执行文件内容读取
                    Task task = Task.Run(() =>
                    {
                        dealExcelAndPrintLogThreadMethod(directoryInfo);
                    });
                }
            }
        }

        //读取Excel中内容并打印日志的方法（传过来的参数是选择目录下的信息对象）
        private void dealExcelAndPrintLogFun(DirectoryInfo directoryInfo)
        {
            //创建汇总文件夹和异常数据文件夹
            string summaryFilePath = directoryInfo.FullName + "\\汇总";
            string errorFilePath = directoryInfo.FullName + "\\异常数据\\";
            if (!Directory.Exists(summaryFilePath))
            {
                Directory.CreateDirectory(summaryFilePath);
            }
            else {
                //删除对应文件下所有文件
                FileOptUtils.DeleteAllFiles(summaryFilePath);
            }
            if (!Directory.Exists(errorFilePath))
            {
                Directory.CreateDirectory(errorFilePath);
            }
            else {
                //删除对应文件下所有文件
                //FileOptUtils.DeleteAllFiles(errorFilePath);
            }

            FileSystemInfo[] fsInfos = directoryInfo.GetFiles();
            //定义对象存储Excel中的数据
            List< CompanyRDData > companyRDDatas = new List< CompanyRDData >();
            foreach (FileSystemInfo fsInfo in fsInfos)
            {
                //跳过隐藏文件
                if (fsInfo.Attributes.HasFlag(FileAttributes.Hidden)) {
                    continue;
                }

                //是否纳入汇总
                Boolean isSummary = true;
                //错误日志信息
                String errorText = "";
                //是否要提示
                Boolean isTips = false;
                //提示日志信息
                String tipsText = "";

                //常规日志输出
                logTextBox.AppendText(fsInfo.Name + "\r\n");
                //判断文件是否还存在，并判断文件类型
                if (fsInfo.Exists && (fsInfo.Extension == ".xls" || fsInfo.Extension == ".xlsx"))
                {
                    CompanyRDData companyRDData = new CompanyRDData();
                    FileStream fs = new FileStream(fsInfo.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    IWorkbook workbook;
                    // 根据文件扩展名判断是.xls还是.xlsx
                    if (fsInfo.Extension == ".xls")
                    {
                        workbook = new HSSFWorkbook(fs);
                    }
                    else if (fsInfo.Extension == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(fs);
                    }
                    else
                    {
                        errorLogRichTextBox.AppendText("不支持的文件：" + fsInfo.Name + "\r\n");
                        continue;
                    }

                    //获取Excel中的数据
                    // 第一个工作表对象（107-1：企业研发项目填报报表）
                    ISheet RDProjectSheet = workbook.GetSheetAt(0);
                    // 第二个工作表对象（107-2：企业研发活动相关情况表）
                    ISheet RDAcitivitySheet = workbook.GetSheetAt(1);
                    //开始获取第一个表格（107-1：企业研发项目填报报表）的数据
                    Console.WriteLine(fsInfo.FullName);
                    Get1071SheetData(companyRDData, RDProjectSheet);
                    //开始获取第二个Sheet表（107-2：企业研发活动相关情况表）的数据
                    Get1072SheetData(companyRDData, RDAcitivitySheet);

                    //查询数据库获取下目前库中已有的项目，用于校验项目名称与往年项目名称的重合度
                    string queryHistoryProjectInfoSql = "SELECT Company_Name,Project_Name,Project_Year FROM Company_History_ProjectName";
                    DataTable historyProjectInfoDataTable = DBHelper.ExecuteQuery(queryHistoryProjectInfoSql);

                    //逻辑判断
                    /**
                      思路：系统分为两个等级：1、提示等级（仅提示，但输入会纳入汇总中）；2、错误等级（显示错误信息，不会纳入系统）
                      因此这两部分的提示信息需要拆开，因为有些填报的数据既触发了错误等级又触发提示等级，这些得分开
                     * */
                    LogicCheck(companyRDData, historyProjectInfoDataTable, ref isSummary, ref errorText, ref isTips, ref tipsText);

                    //可以合并计算的数据
                    if (isSummary)
                    {
                        companyRDDatas.Add(companyRDData);
                    }
                    else {
                        //设置错误类信息文本变色
                        errorLogRichTextBox.SelectionStart = errorLogRichTextBox.Text.Length;
                        errorLogRichTextBox.SelectionLength = 1;
                        errorLogRichTextBox.SelectionBackColor = Color.FromName("Red");

                        errorLogRichTextBox.AppendText("错误提示：《" + fsInfo.Name+"》" + "中违反了以下规则：\r\n" + errorText+ "\r\n");
                        //拷贝一份文件到异常数据文件夹
                        /*if (File.Exists(errorFilePath + fsInfo.Name))
                        {
                            File.Delete(errorFilePath + fsInfo.Name);
                        }
                        else { 
                            File.Move(fsInfo.FullName, errorFilePath + fsInfo.Name);
                        }*/
                    }

                    //常规提示性信息打印
                    if (isTips) {
                        //提醒不剔除数据的
                        errorLogRichTextBox.AppendText("建议完善以下方面：《" + fsInfo.Name + "》" + "中违反了以下规则：\r\n" + tipsText + "\r\n\r\n");
                    }

                    fs.Close();
                }
                else { 
                    errorLogRichTextBox.AppendText("不支持的文件：" + fsInfo.Name + "\r\n");
                }
            }

            //开始遍历并合并数据
            CompanyRDData summaryCompanyRDData = new CompanyRDData();
            summaryDataFun(companyRDDatas, summaryCompanyRDData);

            //导出数据到Excel中
            exportSummaryDataIntoExcel(summaryCompanyRDData, summaryFilePath);

            //把项目人员实际工作时间输出到前台
            logTextBox.AppendText("\r\n该县市区研发填报情况如下：\r\n");
            logTextBox.AppendText("研究开发费用合计：" + Math.Round(summaryCompanyRDData.RD1071ExpensesTotal/ 100000, 4) + "亿元\r\n");
            logTextBox.AppendText("研发人员全时当量合计：" + Math.Round(summaryCompanyRDData.RDProjectStaffWorkMonth / 12, 2) + "人年\r\n");
            logTextBox.AppendText("符合填报要求企业数量合计：" + companyRDDatas.Count + "家\r\n");
        }

        //获取107-1表中的数据并赋值到对象的列表中
        private void Get1071SheetData(CompanyRDData companyRDData, ISheet RDProjectSheet) {
            Console.WriteLine("开始读取107-1数据");

            /*  
            基本思路：1、从6行开始读取数据
                      2、判断每一行的第一个单元格的不是以“单位负责人： ”开头，而且第二个单元格有项目名称
                      3、开始获取数据
            */
            for (int i = 5; i < RDProjectSheet.LastRowNum; i++) { 
                //索引列
                String indexColStr = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 0);
                //项目名称列
                String RDProjectName = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 1);
                //项目
                if (indexColStr.StartsWith("单位负责人") || string.IsNullOrEmpty(RDProjectName)) {
                    break;
                }

                //创建一条项目对象，用来接收项目信息
                ProjectRDData projectRDData = new ProjectRDData();
                //项目名称
                projectRDData.RDProjectName = RDProjectName;
                //项目来源
                projectRDData.RDProjectSource = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 2);
                //项目开展形式
                projectRDData.RDProjectDevForm = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 3);
                //项目当年成果形式(Project current results form)
                projectRDData.RDProjectCurrentResultsForm = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 4);
                //项目技术经济目标
                projectRDData.RDProjectEconomicTarget = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 5);
                //项目起始日期
                projectRDData.RDProjectBeginDate = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 6);
                //项目完成日期
                projectRDData.RDProjectEndDate = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 7);
                //跨年项目当年所处主要进展阶段
                projectRDData.AcrossYearRDProjectCurrentStage = ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 8);
                //项目研究开发人员 （人）
                projectRDData.RDProjectResearcherCount = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 9));
                //项目人员实际工作时间  （人月）
                projectRDData.RDProjectStaffWorkMonth = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 10));
                //项目经费支出（千元）
                projectRDData.RDProjectExpenses = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 11));
                //其中：政府资金
                projectRDData.RDProjectExpensesFromGovernment = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 12));
                //*其中：用于科学原理的探索发现
                projectRDData.RDProjectExpensesForSicResearch = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 13));
                //*其中：企业自主开展
                projectRDData.RDProjectExpensesFromComSelf = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 14));
                //*委托外单位开展
                projectRDData.RDProjectExpensesFromEntrustOutsource = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDProjectSheet, i, 15));

                //把项目信息对象塞入到企业信息表中
                companyRDData.projectRDDatas.Add(projectRDData);
            }
        }


        //获取107-2表中的数据并赋值到对象中
        private void Get1072SheetData(CompanyRDData companyRDData, ISheet RDAcitivitySheet)
        {
            //研究开发人员合计(人)
            companyRDData.RDPersonnelTotal = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 5, 3));
            //其中：管理和服务人员
            companyRDData.RDPersonnelManageAndService = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 6, 3));
            //其中：女性(人)
            companyRDData.RDPersonnelFemale = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 7, 3));
            //其中：全职人员
            companyRDData.RDPersonnelFullTimeStaff = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 8, 3));
            //其中：本科毕业及以上人员(人)
            companyRDData.RDPersonnelBachelorAndAbove = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 9, 3));
            //其中：外聘人员(人)
            companyRDData.RDPersonnelExternalStaff = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 10, 3));

            //二、研究开发费用情况
            // 研究开发费用合计（千元）
            companyRDData.RDExpensesTotal = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 12, 3));
            //1.人员人工费用（千元）
            companyRDData.RDExpensesPersonnelLabor = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 13, 3));
            //2.直接投入费用（千元）
            companyRDData.RDExpensesDirectInput = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 14, 3));
            //3.折旧费用与长期待摊费用（千元）
            companyRDData.RDExpensesDepreciationAndLongTerm = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 15, 3));
            //4.无形资产摊销费用（千元）
            companyRDData.RDExpensesIntangibleAssets = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 16, 3));
            //5.设计费用（千元）
            companyRDData.RDExpensesDesign = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 17, 3));
            //6.装备调试费用与试验费用（千元）
            companyRDData.RDExpensesEquipmentDebug = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 18, 3));
            //7.委托外部研究开发费用（千元）
            companyRDData.RDExpensesEntrustOutsourcedRD = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 19, 3));
            //①委托境内研究机构（千元）
            companyRDData.RDExpensesEntrustDomesticResearch = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 20, 3));
            //②委托境内高等学校（千元）
            companyRDData.RDExpensesEntrustDomesticCollege = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 21, 3));
            //③委托境内企业（千元）
            companyRDData.RDExpensesEntrustDomesticCompany = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 22, 3));
            //④委托境外机构（千元）
            companyRDData.RDExpensesEntrustOverseasInstitutions = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 23, 3));
            //8.其他费用（千元）
            companyRDData.RDExpensesOthers = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 24, 3));

            //三、研究开发资产情况
            //当年形成用于研究开发的固定资产（千元）
            companyRDData.RDAssetsYear = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 26, 3));
            //其中：仪器和设备（千元）
            companyRDData.RDAssetsYearEquipment = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 27, 3));

            //四、研究开发支出资金来源
            //1.来自企业自筹(千元)
            companyRDData.RDSpendSourceOfCompany = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 29, 3));
            //2.来自政府部门（千元）
            companyRDData.RDSpendSourceOfGovernment = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 30, 3));
            //3.来自银行贷款（千元）
            companyRDData.RDSpendSourceOfBank = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 31, 3));
            //4.来自风险投资（千元）
            companyRDData.RDSpendSourceOfRiskCapital = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 32, 3));
            //5.来自其他渠道（千元）
            companyRDData.RDSpendSourceOfOthers = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 33, 3));

            //五、相关政策落实情况
            //申报加计扣除减免税的研究开发支出(千元)
            companyRDData.PolicyImplementDeclareAddtionRD = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 35, 3));
            //加计扣除减免税金额(千元)
            companyRDData.PolicyImplementAddtionRDTaxFree = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 36, 3));
            //高新技术企业减免税金额(千元)
            companyRDData.PolicyImplementHighTechRDTaxFree = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 37, 3));

            //六、企业办研究开发机构（境内）情况
            //期末机构数(个)
            companyRDData.CompanyRunOrgCountEndOfPeriod = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 5, 10));
            //机构研究开发人员（人）
            companyRDData.CompanyRunOrgRDPersonnel = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 6, 10));
            //其中：博士毕业（人）
            companyRDData.CompanyRunOrgRDDoctor = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 7, 10));
            //其中：硕士毕业（人）
            companyRDData.CompanyRunOrgRDMaster = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 8, 10));
            //机构研究开发费用（千元）
            companyRDData.CompanyRunOrgRDExpenses = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 9, 10));
            //期末仪器和设备原价（千元）
            companyRDData.CompanyRunOrgEquipmentValueEndOfPeriod = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 10, 10));

            //七、研究开发产出及相关情况
            //(一) 专利情况
            //当年专利申请数(件)
            companyRDData.PatentApplyOfCurrentYear = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 13, 10));
            //其中：发明专利（件）
            companyRDData.PatentApplyOfInvention = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 14, 10));
            //期末有效发明专利数（件）
            companyRDData.PatentApplyOfInForcePeriod = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 15, 10));
            //其中：已被实施（件）
            companyRDData.PatentApplyOfBeenImplement = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 16, 10));
            //专利所有权转让及许可数（件）
            companyRDData.PatentApplyOfAssignment = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 17, 10));
            //专利所有权转让及许可收入（千元）
            companyRDData.PatentApplyOfAssignmentIncome = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 18, 10));
            //(二) 新产品情况
            //*新产品销售收入(千元)
            companyRDData.NewProductSaleRevenue = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 20, 10));
            //*其中：出口(千元)
            companyRDData.NewProductSaleOfOutlet = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 21, 10));
            //(三)其他情况
            //*期末拥有注册商标(件)
            companyRDData.TrademarkOfPeriod = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 23, 10));
            //发表科技论文(篇)
            companyRDData.ScientificPapers = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 24, 10));
            //形成国家或行业标准(项)
            companyRDData.StandardsOfNational = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 25, 10));

            //八、其他相关情况
            //(一)技术改造和技术获取情况
            //技术改造经费支出（千元）
            companyRDData.TechTransformExpenses = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 28, 10));
            //购买境内技术经费支出（千元）
            companyRDData.BuyDomesticTechExpenses = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 29, 10));
            //引进境外技术经费支出（千元）
            companyRDData.ImpOverseasTechExpenses = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 30, 10));
            //引进境外技术的消化吸收经费支出（千元）
            companyRDData.ImpOverseasTechDigestionExpenses = NumberUtils.getDecimal(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 31, 10));
            // (二)企业办研究开发机构（境外）情况
            //期末企业在境外设立的研究开发机构数(个)
            companyRDData.OverseasOrgCount = NumberUtils.getInt(ExcelUtils.getCellValueByCellType(RDAcitivitySheet, 33, 10));
        }

        //逻辑校验
        private void LogicCheck(CompanyRDData companyRDData, DataTable historyProjectInfoDataTable, ref Boolean isSummary, ref String errorText,ref Boolean isTips, ref String tipsText) {

            //107-2表 企业研究开发活动及相关情况
            List<ProjectRDData> projectRDDatas = companyRDData.projectRDDatas;
            //下面这些是将数据剔除出去的条件
            //1≥2（研究开发人员合计≥其中：管理和服务人员）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelManageAndService)
            {
                isTips = true;
                tipsText += "研究开发人员合计≥其中：管理和服务人员；\r\n";
            }
            //1≥3（研究开发人员合计≥其中：女性）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelFemale)
            {
                isTips = true;
                tipsText += "研究开发人员合计≥其中：女性；\r\n";
            }
            //1≥4（研究开发人员合计≥其中：全职人员）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelFullTimeStaff)
            {
                isTips = true;
                tipsText += "研究开发人员合计≥其中：全职人员；\r\n";
            }
            //1≥5≥24+25（研究开发人员合计≥其中：本科毕业及以上人员≥其中：博士毕业+其中：硕士毕业）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelBachelorAndAbove || companyRDData.RDPersonnelBachelorAndAbove < (companyRDData.CompanyRunOrgRDDoctor + companyRDData.CompanyRunOrgRDMaster))
            {
                isTips = true;
                tipsText += "研究开发人员合计≥其中：本科毕业及以上人员≥其中：博士毕业+其中：硕士毕业；\r\n";
            }
            //1≥6（研究开发人员合计≥其中：外聘人员）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelExternalStaff)
            {
                isTips = true;
                tipsText += "研究开发人员合计≥其中：外聘人员；\r\n";
            }
            //1≥23≥24+25（研究开发人员合计≥机构研究开发人员≥其中：博士毕业+其中：硕士毕业）
            if (companyRDData.RDPersonnelTotal < companyRDData.CompanyRunOrgRDPersonnel || companyRDData.CompanyRunOrgRDPersonnel < (companyRDData.CompanyRunOrgRDDoctor + companyRDData.CompanyRunOrgRDMaster))
            {
                isTips = true;
                tipsText += "研究开发人员合计≥机构研究开发人员≥其中：博士毕业+其中：硕士毕业；\r\n";
            }
            //7=8+9+10+11+12+13+14+19≥26（研究开发费用合计=人员人工费用+直接投入费用+折旧费用与长期待摊费用+无形资产摊销费用+设计费用+装备调试费用与试验费用+委托外部研究开发费用+其他费用≥机构研究开发费用）
            if (companyRDData.RDExpensesTotal != (companyRDData.RDExpensesPersonnelLabor + companyRDData.RDExpensesDirectInput + companyRDData.RDExpensesDepreciationAndLongTerm + companyRDData.RDExpensesIntangibleAssets + companyRDData.RDExpensesDesign + companyRDData.RDExpensesEquipmentDebug + companyRDData.RDExpensesEntrustOutsourcedRD + companyRDData.RDExpensesOthers) || companyRDData.RDExpensesTotal < companyRDData.CompanyRunOrgRDExpenses)
            {
                isTips = true;
                tipsText += "研究开发费用合计=人员人工费用+直接投入费用+折旧费用与长期待摊费用+无形资产摊销费用+设计费用+装备调试费用与试验费用+委托外部研究开发费用+其他费用≥机构研究开发费用；\r\n";
            }
            //若1>0，则8>0（若研究开发人员合计>0，则人员人工费用>0）
            if (companyRDData.RDPersonnelTotal > 0 && companyRDData.RDExpensesPersonnelLabor == 0)
            {
                isTips = true;
                tipsText += "若研究开发人员合计>0，则人员人工费用>0；\r\n";
            }
            //若8>0，则1>0（若人员人工费用>0，则研究开发人员合计>0）
            if (companyRDData.RDExpensesPersonnelLabor > 0 && companyRDData.RDPersonnelTotal == 0)
            {
                isTips = true;
                tipsText += "若人员人工费用>0，则研究开发人员合计>0；\r\n";
            }
            //14=15+16+17+18（委托外部研究开发费用=①委托境内研究机构+②委托境内高等学校+③委托境内企业+④委托境外机构）
            if (companyRDData.RDExpensesEntrustOutsourcedRD != (companyRDData.RDExpensesEntrustDomesticResearch + companyRDData.RDExpensesEntrustDomesticCollege + companyRDData.RDExpensesEntrustDomesticCompany + companyRDData.RDExpensesEntrustOverseasInstitutions))
            {
                isTips = true;
                tipsText += "委托外部研究开发费用=①委托境内研究机构+②委托境内高等学校+③委托境内企业+④委托境外机构；\r\n";
            }
            //20≥21（当年形成用于研究开发的固定资产≥其中：仪器和设备）
            if (companyRDData.RDAssetsYear < companyRDData.RDAssetsYearEquipment)
            {
                isTips = true;
                tipsText += "当年形成用于研究开发的固定资产≥其中：仪器和设备；\r\n";
            }
            //若27>0，则22>0（若期末仪器和设备原价>0，则期末机构数>0）
            if (companyRDData.CompanyRunOrgEquipmentValueEndOfPeriod > 0 && companyRDData.CompanyRunOrgCountEndOfPeriod == 0)
            {
                isTips = true;
                tipsText += "若期末仪器和设备原价>0，则期末机构数>0；\r\n";
            }
            //29≥30（当年专利申请数≥其中：发明专利）
            if (companyRDData.PatentApplyOfCurrentYear < companyRDData.PatentApplyOfInvention)
            {
                isTips = true;
                tipsText += "当年专利申请数≥其中：发明专利；\r\n";
            }
            //32≥33（期末有效发明专利数≥其中：已被实施）
            if (companyRDData.PatentApplyOfInForcePeriod < companyRDData.PatentApplyOfBeenImplement)
            {
                isTips = true;
                tipsText += "期末有效发明专利数≥其中：已被实施；\r\n";
            }
            //36≥37（新产品销售收入≥其中：出口）
            if (companyRDData.NewProductSaleRevenue < companyRDData.NewProductSaleOfOutlet)
            {
                isTips = true;
                tipsText += "新产品销售收入≥其中：出口；\r\n";
            }
            //研究开发费用合计 = 四、研究开发支出资金来源中各项的和
            if (companyRDData.RDExpensesTotal != (companyRDData.RDSpendSourceOfCompany + companyRDData.RDSpendSourceOfGovernment + companyRDData.RDSpendSourceOfBank + companyRDData.RDSpendSourceOfRiskCapital + companyRDData.RDSpendSourceOfOthers))
            {
                isTips = true;
                tipsText += "研究开发费用合计 != 四、研究开发支出资金来源中各项的和；\r\n";
            }
            //107-2表中 研究开发人员合计*12需大于各项目人月合计
            if (companyRDData.RDPersonnelTotal * 12 < companyRDData.RDProjectStaffWorkMonth) {
                isTips = true;
                tipsText += "107-2表中 研究开发人员合计*12需大于各项目人月合计；\r\n";
            }

            //下面这些是只做提醒的的条件
            //"期末机构数"如果为0，则进行提醒，不剔除数据
            if (companyRDData.CompanyRunOrgCountEndOfPeriod == 0)
            {
                isTips = true;
                tipsText += "期末机构数为0；\r\n";
            }
            //" 7.委托外部研究开发费用"如果大于0，，则进行提醒，不剔除数据
            if (companyRDData.RDExpensesEntrustOutsourcedRD > 0)
            {
                isTips = true;
                tipsText += "存在委托外部研究开发费用；\r\n";
            }

            //107-2申报加计扣除为0的提示出来
            if (companyRDData.PolicyImplementAddtionRDTaxFree == 0)
            {
                isTips = true;
                tipsText += "加计扣除减免税金额为0千元；\r\n";
            }
            //当年专利申请为0或发表的科技论文数为0的提示
            if (companyRDData.PatentApplyOfCurrentYear == 0 && companyRDData.ScientificPapers == 0)
            {
                isTips = true;
                tipsText += "当年专利申请数和发表科技论文数为不能同时0；\r\n";
            }

            //校验107-1 项目信息表中的信息
            //校验项目名称中不能包含“一种、技改、改造、年产、生产线、打样、翻样、产业化、示范、推广、CYH、JG、系统”字样
            String[] forbiddenWordsInProjectNameArg = { "土建", "产业园", "资产", "购买", "改良", "一种", "改造", "技改", "生产", "产业化", "打样", "翻样", "推广", "示范", "年产", "业务费", "运行补贴", "补贴", "工业性实验", "成果应用", "成果转化", "绿色制造", "节能", "产业", "国债付息", "创投", "创业投资", "科技特派员", "特派员", "技术服务", "技术创新服务", "智库", "科普", "科学技术普及", "运营维护", "升级", "科技交流", "智慧信访", "工作经费", "信息系统", "信息化平台", "办公平台", "平台建设", "管理平台", "管理系统", "监管平台", "智慧校园", "离退休", "业务活动", "奖励", "股改", "融资", "租赁", "科技管理", "备案", "评估", "服务经费", "项目监理", "打包贷款", "技术改造", "生产线", "CYH", "JG", "系统", "升级改造" };

            for (int i = projectRDDatas.Count-1; i >=0 ; i--) {
                ProjectRDData projectRDData = projectRDDatas[i];

                //是否移除项目（2024-4-17 姜昊科长让把107-1表中不符合要求的项目不统计在内）
                Boolean isRemoveProject = false;
                //项目名称规则校验
                String projectName = projectRDData.RDProjectName;
                String forbiddenWords = "";
                foreach (String forbiddenWordInProjectName in forbiddenWordsInProjectNameArg)
                {
                    if (projectName.Contains(forbiddenWordInProjectName)) {
                        forbiddenWords += (forbiddenWordInProjectName+"、");
                    }
                }
                if (!string.IsNullOrEmpty(forbiddenWords)) {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目名称中包含\""+ forbiddenWords.Substring(0, forbiddenWords.Length - 1) + "\"字眼，该项目不计入研发统计数据；\r\n");
                }

                //校验项目名称中包含连续字母的提示
                String mentionEngCharStr = StringUtils.getContainsChar(projectName);
                if (mentionEngCharStr.Length >= 3) {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目名称中包含\"" + mentionEngCharStr + "\"连续英文字母，该项目不计入研发统计数据；\r\n");
                }

                //校验项目名称是否与往年的项目存在高重合度




                //项目当年成果形式中如果包含了（2.新产品、新工艺等推广与示范活动或3.对已有产品、工艺等进行一般性改进）则进行提示
                if (projectRDData.RDProjectCurrentResultsForm.StartsWith("2")) {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目当年成果形式不能选择2.新产品、新工艺等推广与示范活动，该项目不计入研发统计数据；\r\n");
                }else if (projectRDData.RDProjectCurrentResultsForm.StartsWith("3"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目当年成果形式不能选择3.对已有产品、工艺等进行一般性改进，该项目不计入研发统计数据；\r\n");
                }else if (projectRDData.RDProjectCurrentResultsForm.StartsWith("11"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目当年成果形式不能选择11.带有技术、工艺参数的图纸、技术标准、操作规范、技术论证、咨询评价，该项目不计入研发统计数据；\r\n");
                }else if (projectRDData.RDProjectCurrentResultsForm.StartsWith("14"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目当年成果形式不能选择14.其他形式，该项目不计入研发统计数据；\r\n");
                }

                //技术经济指标选5.提高劳动生产率、6.减少能源消耗或提高能源使用效率、7.节约原材料、8.减少环境污染（提示）
                if (projectRDData.RDProjectEconomicTarget.StartsWith("5"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目技术经济目标不能选择5.提高劳动生产率，该项目不计入研发统计数据；\r\n");
                }
                if (projectRDData.RDProjectEconomicTarget.StartsWith("6"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目技术经济目标不能选择6.减少能源消耗或提高能源使用效率，该项目不计入研发统计数据；\r\n");
                }
                if (projectRDData.RDProjectEconomicTarget.StartsWith("7"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目技术经济目标不能选择7.节约原材料，该项目不计入研发统计数据；\r\n");
                }
                if (projectRDData.RDProjectEconomicTarget.StartsWith("8"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目技术经济目标不能选择8.减少环境污染，该项目不计入研发统计数据；\r\n");
                }
                if (projectRDData.RDProjectEconomicTarget.StartsWith("9"))
                {
                    isTips = true;
                    isRemoveProject = true;
                    tipsText += (projectName + "项目技术经济目标不能选择9.其他，该项目不计入研发统计数据；\r\n");
                }

                /*  //项目起始日期&项目完成日期
                  DateTime ProjectBeginDate = DateUtils.formatDatetime(projectRDData.RDProjectBeginDate);
                  DateTime ProjectEndDate = DateUtils.formatDatetime(projectRDData.RDProjectEndDate);
                  //如果是项目早于2000年或者项目晚于2045年，则认为是解析日期的时候解析出错了
                  if (ProjectBeginDate.Year < 2000 || ProjectBeginDate.Year > 2045 || ProjectEndDate.Year < 2000 || ProjectEndDate.Year > 2045)
                  {
                      isTips = true;
                      isRemoveProject = true;
                      tipsText += (projectName + "项目的起始日期或项目的完成日期未按照6位格式（202312）填报，该项目不计入研发统计数据；\r\n");
                  }
                  else {
                      //如果项目周期跨年
                      if (ProjectBeginDate.Year != DateTime.Now.Year || ProjectEndDate.Year != DateTime.Now.Year)
                      {
                          //如果“跨年项目需要填写主要进展阶段”为空（跨年项目需要填写跨年项目需要填写主要进展阶段）
                          if (string.IsNullOrEmpty(projectRDData.AcrossYearRDProjectCurrentStage))
                          {
                              isTips = true;
                              isRemoveProject = true;
                              tipsText += (projectName + "项目是跨年项目，但是未选择跨年项目当年所处主要进展阶段，该项目不计入研发统计数据；\r\n");
                          }
                          else {
                              //如果不是选择1.研究阶段和2.小试阶段开头的都需要排除
                              if (!(projectRDData.AcrossYearRDProjectCurrentStage.StartsWith("1") || projectRDData.AcrossYearRDProjectCurrentStage.StartsWith("2"))) {
                                  isTips = true;
                                  isRemoveProject = true;
                                  tipsText += (projectName + "项目是跨年项目，跨年项目当年所处主要进展阶段选择了非“1.研究阶段和2.小试阶段”，该项目不计入研发统计数据；\r\n");
                              }
                          }
                      }

                      //项目周期要大约等于3个月
                      if (ProjectEndDate.Month - ProjectBeginDate.Month + (ProjectEndDate.Year - ProjectBeginDate.Year) * 12 +1< 4)
                      {
                          isTips = true;
                          //isRemoveProject = true;
                          tipsText += (projectName + "项目的周期必须大于3个月；\r\n");
                      }
                  }*/

                //如果107-1表中项目有不符合规则的，则从项目中删除，不纳入最后的研发费用合计
                if (isRemoveProject)
                {
                    companyRDData.projectRDDatas.RemoveAt(i);
                }
                else {
                    //如果不移除，说明数据合法
                    //计算下人月合计最后赋值到companyRDData对象中，供后期计算人月工资使用
                    companyRDData.RDProjectStaffWorkMonth += projectRDData.RDProjectStaffWorkMonth;
                    //计算107-1表中所有项目的研发投入的合计，供后面107-2表中研发投入合计使用
                    companyRDData.RD1071ExpensesTotal += projectRDData.RDProjectExpenses;
                }
            }
            Console.WriteLine("研发人员全时当量："+companyRDData.RDProjectStaffWorkMonth);

            //人员费用支出/人月,低于<2200，不能高于5万（提示）
            if (companyRDData.RDProjectStaffWorkMonth == 0)
            {
                isTips = true;
                tipsText += "107-1表：项目人员实际工作时间为0；\r\n";
            }
            else
            {
                decimal avgWagesPerMonth = companyRDData.RDExpensesPersonnelLabor * 1000 / companyRDData.RDProjectStaffWorkMonth;
                if (avgWagesPerMonth < 2200 || avgWagesPerMonth > 50000)
                {
                    isTips = true;
                    tipsText += "107-2表：人员人工费用÷107-1表：项目人员实际工作时间（人月）合计小于2200元或大于5万元；\r\n";
                }
            }

            //如果这家企业一个合法的项目都没有，直接把这家企业剔除掉
            if (companyRDData.projectRDDatas.Count == 0)
            {
                isSummary = false;
                errorText = "所有项目都不满足统计要求，该企业数据将不被统计在内！";
            }
        }

        //汇总数据的和
        private void summaryDataFun(List<CompanyRDData> companyRDDatas, CompanyRDData summaryCompanyRDData) {
            foreach (var companyRDData in companyRDDatas)
            {
                //研究开发人员合计(人)
                summaryCompanyRDData.RDPersonnelTotal += companyRDData.RDPersonnelTotal;
                //其中：管理和服务人员
                summaryCompanyRDData.RDPersonnelManageAndService += companyRDData.RDPersonnelManageAndService;
                //其中：女性(人)
                summaryCompanyRDData.RDPersonnelFemale += companyRDData.RDPersonnelFemale;
                //其中：全职人员
                summaryCompanyRDData.RDPersonnelFullTimeStaff += companyRDData.RDPersonnelFullTimeStaff;
                //其中：本科毕业及以上人员(人)
                summaryCompanyRDData.RDPersonnelBachelorAndAbove += companyRDData.RDPersonnelBachelorAndAbove;
                //其中：外聘人员(人)
                summaryCompanyRDData.RDPersonnelExternalStaff += companyRDData.RDPersonnelExternalStaff;

                //二、研究开发费用情况
                // 研究开发费用合计（千元）
                //summaryCompanyRDData.RDExpensesTotal = summaryCompanyRDData.RDExpensesTotal + companyRDData.RDExpensesTotal;
                // 2024-4-17 姜昊科长指示，研究开发费用合计使用所有和项目的研发费用和计算
                summaryCompanyRDData.RDExpensesTotal = summaryCompanyRDData.RDExpensesTotal + companyRDData.RDExpensesTotal;
                //1.人员人工费用（千元）
                summaryCompanyRDData.RDExpensesPersonnelLabor = summaryCompanyRDData.RDExpensesPersonnelLabor + companyRDData.RDExpensesPersonnelLabor;
                //2.直接投入费用（千元）
                summaryCompanyRDData.RDExpensesDirectInput = summaryCompanyRDData.RDExpensesDirectInput + companyRDData.RDExpensesDirectInput;
                //3.折旧费用与长期待摊费用（千元）
                summaryCompanyRDData.RDExpensesDepreciationAndLongTerm = summaryCompanyRDData.RDExpensesDepreciationAndLongTerm + companyRDData.RDExpensesDepreciationAndLongTerm;
                //4.无形资产摊销费用（千元）
                summaryCompanyRDData.RDExpensesIntangibleAssets = summaryCompanyRDData.RDExpensesIntangibleAssets + companyRDData.RDExpensesIntangibleAssets;

                //5.设计费用（千元）
                summaryCompanyRDData.RDExpensesDesign = summaryCompanyRDData.RDExpensesDesign + companyRDData.RDExpensesDesign;

                //6.装备调试费用与试验费用（千元）
                summaryCompanyRDData.RDExpensesEquipmentDebug = summaryCompanyRDData.RDExpensesEquipmentDebug + companyRDData.RDExpensesEquipmentDebug;
                //7.委托外部研究开发费用（千元）
                summaryCompanyRDData.RDExpensesEntrustOutsourcedRD = summaryCompanyRDData.RDExpensesEntrustOutsourcedRD + companyRDData.RDExpensesEntrustOutsourcedRD;
                //①委托境内研究机构（千元）
                summaryCompanyRDData.RDExpensesEntrustDomesticResearch = summaryCompanyRDData.RDExpensesEntrustDomesticResearch + companyRDData.RDExpensesEntrustDomesticResearch;
                //②委托境内高等学校（千元）
                summaryCompanyRDData.RDExpensesEntrustDomesticCollege = summaryCompanyRDData.RDExpensesEntrustDomesticCollege + companyRDData.RDExpensesEntrustDomesticCollege;
                //③委托境内企业（千元）
                summaryCompanyRDData.RDExpensesEntrustDomesticCompany = summaryCompanyRDData.RDExpensesEntrustDomesticCompany + companyRDData.RDExpensesEntrustDomesticCompany;
                //④委托境外机构（千元）
                summaryCompanyRDData.RDExpensesEntrustOverseasInstitutions = summaryCompanyRDData.RDExpensesEntrustOverseasInstitutions + companyRDData.RDExpensesEntrustOverseasInstitutions;
                //8.其他费用（千元）
                summaryCompanyRDData.RDExpensesOthers = summaryCompanyRDData.RDExpensesOthers + companyRDData.RDExpensesOthers;

                //三、研究开发资产情况
                //当年形成用于研究开发的固定资产（千元）
                summaryCompanyRDData.RDAssetsYear = summaryCompanyRDData.RDAssetsYear + companyRDData.RDAssetsYear;
                //其中：仪器和设备（千元）
                summaryCompanyRDData.RDAssetsYearEquipment = summaryCompanyRDData.RDAssetsYearEquipment + companyRDData.RDAssetsYearEquipment;


                //四、研究开发支出资金来源
                //1.来自企业自筹(千元)
                summaryCompanyRDData.RDSpendSourceOfCompany = summaryCompanyRDData.RDSpendSourceOfCompany + companyRDData.RDSpendSourceOfCompany;
                //2.来自政府部门（千元）
                summaryCompanyRDData.RDSpendSourceOfGovernment = summaryCompanyRDData.RDSpendSourceOfGovernment + companyRDData.RDSpendSourceOfGovernment;
                //3.来自银行贷款（千元）
                summaryCompanyRDData.RDSpendSourceOfBank = summaryCompanyRDData.RDSpendSourceOfBank + companyRDData.RDSpendSourceOfBank;
                //4.来自风险投资（千元）
                summaryCompanyRDData.RDSpendSourceOfRiskCapital = summaryCompanyRDData.RDSpendSourceOfRiskCapital + companyRDData.RDSpendSourceOfRiskCapital;
                //5.来自其他渠道（千元）
                summaryCompanyRDData.RDSpendSourceOfOthers = summaryCompanyRDData.RDSpendSourceOfOthers + companyRDData.RDSpendSourceOfOthers;

                //五、相关政策落实情况
                //申报加计扣除减免税的研究开发支出(千元)
                summaryCompanyRDData.PolicyImplementDeclareAddtionRD = summaryCompanyRDData.PolicyImplementDeclareAddtionRD + companyRDData.PolicyImplementDeclareAddtionRD;
                //加计扣除减免税金额(千元)
                summaryCompanyRDData.PolicyImplementAddtionRDTaxFree = summaryCompanyRDData.PolicyImplementAddtionRDTaxFree + companyRDData.PolicyImplementAddtionRDTaxFree;
                //高新技术企业减免税金额(千元)
                summaryCompanyRDData.PolicyImplementHighTechRDTaxFree = summaryCompanyRDData.PolicyImplementHighTechRDTaxFree + companyRDData.PolicyImplementHighTechRDTaxFree;


                //六、企业办研究开发机构（境内）情况
                //期末机构数(个)
                summaryCompanyRDData.CompanyRunOrgCountEndOfPeriod += companyRDData.CompanyRunOrgCountEndOfPeriod;
                //机构研究开发人员（人）
                summaryCompanyRDData.CompanyRunOrgRDPersonnel += companyRDData.CompanyRunOrgRDPersonnel;
                //其中：博士毕业（人）
                summaryCompanyRDData.CompanyRunOrgRDDoctor += companyRDData.CompanyRunOrgRDDoctor;
                //其中：硕士毕业（人）
                summaryCompanyRDData.CompanyRunOrgRDMaster += companyRDData.CompanyRunOrgRDMaster;
                //机构研究开发费用（千元）
                summaryCompanyRDData.CompanyRunOrgRDExpenses = summaryCompanyRDData.CompanyRunOrgRDExpenses + companyRDData.CompanyRunOrgRDExpenses;
                //期末仪器和设备原价（千元）
                summaryCompanyRDData.CompanyRunOrgEquipmentValueEndOfPeriod = summaryCompanyRDData.CompanyRunOrgEquipmentValueEndOfPeriod + companyRDData.CompanyRunOrgEquipmentValueEndOfPeriod;

                //七、研究开发产出及相关情况
                //(一) 专利情况
                //当年专利申请数(件)
                summaryCompanyRDData.PatentApplyOfCurrentYear += companyRDData.PatentApplyOfCurrentYear;
                //其中：发明专利（件）
                summaryCompanyRDData.PatentApplyOfInvention += companyRDData.PatentApplyOfInvention;
                //期末有效发明专利数（件）
                summaryCompanyRDData.PatentApplyOfInForcePeriod += companyRDData.PatentApplyOfInForcePeriod;
                //其中：已被实施（件）
                summaryCompanyRDData.PatentApplyOfBeenImplement += companyRDData.PatentApplyOfBeenImplement;
                //专利所有权转让及许可数（件）
                summaryCompanyRDData.PatentApplyOfAssignment += companyRDData.PatentApplyOfAssignment;
                //专利所有权转让及许可收入（千元）
                summaryCompanyRDData.PatentApplyOfAssignmentIncome = summaryCompanyRDData.PatentApplyOfAssignmentIncome + companyRDData.PatentApplyOfAssignmentIncome;

                //(二) 新产品情况
                //*新产品销售收入(千元)
                summaryCompanyRDData.NewProductSaleRevenue = summaryCompanyRDData.NewProductSaleRevenue + companyRDData.NewProductSaleRevenue;
                //*其中：出口(千元)
                summaryCompanyRDData.NewProductSaleOfOutlet = summaryCompanyRDData.NewProductSaleOfOutlet + companyRDData.NewProductSaleOfOutlet;
                //(三)其他情况
                //*期末拥有注册商标(件)
                summaryCompanyRDData.TrademarkOfPeriod += companyRDData.TrademarkOfPeriod;
                //发表科技论文(篇)
                summaryCompanyRDData.ScientificPapers += companyRDData.ScientificPapers;
                //形成国家或行业标准(项)
                summaryCompanyRDData.StandardsOfNational += companyRDData.StandardsOfNational;

                //八、其他相关情况
                //(一)技术改造和技术获取情况
                //技术改造经费支出（千元）
                summaryCompanyRDData.TechTransformExpenses = summaryCompanyRDData.TechTransformExpenses + companyRDData.TechTransformExpenses;
                //购买境内技术经费支出（千元）
                summaryCompanyRDData.BuyDomesticTechExpenses = summaryCompanyRDData.BuyDomesticTechExpenses + companyRDData.BuyDomesticTechExpenses;
                //引进境外技术经费支出（千元）
                summaryCompanyRDData.ImpOverseasTechExpenses = summaryCompanyRDData.ImpOverseasTechExpenses + companyRDData.ImpOverseasTechExpenses;
                //引进境外技术的消化吸收经费支出（千元）
                summaryCompanyRDData.ImpOverseasTechDigestionExpenses = summaryCompanyRDData.ImpOverseasTechDigestionExpenses + companyRDData.ImpOverseasTechDigestionExpenses;
                // (二)企业办研究开发机构（境外）情况
                //期末企业在境外设立的研究开发机构数(个)
                summaryCompanyRDData.OverseasOrgCount += companyRDData.OverseasOrgCount;

                //计算该县市区下所有企业所有项目人员实际工作时间
                summaryCompanyRDData.RDProjectStaffWorkMonth += companyRDData.RDProjectStaffWorkMonth;
                //计算该县市区下所有企业所有项目的研发投入额
                summaryCompanyRDData.RD1071ExpensesTotal += companyRDData.RD1071ExpensesTotal;
            }
        }

        //导出数据到Excel中
        private void exportSummaryDataIntoExcel(CompanyRDData summaryCompanyRDData, String summaryFilePath) {
            //导出Excel
            FileStream summaryFs = new FileStream("EnterprisesRDTemplate.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook summaryWorkbook = new XSSFWorkbook(summaryFs);
            ISheet summarySheet = summaryWorkbook.GetSheetAt(1);
            //填写数据到导出的Excel中
            //研究开发人员合计(人)
            ExcelUtils.writeDataIntoCell(summarySheet, 5, 3, summaryCompanyRDData.RDPersonnelTotal);
            //其中：管理和服务人员
            ExcelUtils.writeDataIntoCell(summarySheet, 6, 3, summaryCompanyRDData.RDPersonnelManageAndService);
            //其中：女性(人)
            ExcelUtils.writeDataIntoCell(summarySheet, 7, 3, summaryCompanyRDData.RDPersonnelFemale);
            //其中：全职人员
            ExcelUtils.writeDataIntoCell(summarySheet, 8, 3, summaryCompanyRDData.RDPersonnelFullTimeStaff);
            //其中：本科毕业及以上人员(人)
            ExcelUtils.writeDataIntoCell(summarySheet, 9, 3, summaryCompanyRDData.RDPersonnelBachelorAndAbove);
            //其中：外聘人员(人)
            ExcelUtils.writeDataIntoCell(summarySheet, 10, 3, summaryCompanyRDData.RDPersonnelExternalStaff);

            //二、研究开发费用情况
            // 研究开发费用合计（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 12, 3, summaryCompanyRDData.RDExpensesTotal);
            //1.人员人工费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 13, 3, summaryCompanyRDData.RDExpensesPersonnelLabor);
            //2.直接投入费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 14, 3, summaryCompanyRDData.RDExpensesDirectInput);
            //3.折旧费用与长期待摊费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 15, 3, summaryCompanyRDData.RDExpensesDepreciationAndLongTerm);
            //4.无形资产摊销费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 16, 3, summaryCompanyRDData.RDExpensesIntangibleAssets);
            //5.设计费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 17, 3, summaryCompanyRDData.RDExpensesDesign);
            //6.装备调试费用与试验费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 18, 3, summaryCompanyRDData.RDExpensesEquipmentDebug);
            //7.委托外部研究开发费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 19, 3, summaryCompanyRDData.RDExpensesEntrustOutsourcedRD);
            //①委托境内研究机构（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 20, 3, summaryCompanyRDData.RDExpensesEntrustDomesticResearch);
            //②委托境内高等学校（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 21, 3, summaryCompanyRDData.RDExpensesEntrustDomesticCollege);
            //③委托境内企业（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 22, 3, summaryCompanyRDData.RDExpensesEntrustDomesticCompany);
            //④委托境外机构（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 23, 3, summaryCompanyRDData.RDExpensesEntrustOverseasInstitutions);
            //8.其他费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 24, 3, summaryCompanyRDData.RDExpensesOthers);

            //三、研究开发资产情况
            //当年形成用于研究开发的固定资产（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 26, 3, summaryCompanyRDData.RDAssetsYear);
            //其中：仪器和设备（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 27, 3, summaryCompanyRDData.RDAssetsYearEquipment);

            //四、研究开发支出资金来源
            //1.来自企业自筹(千元)
            ExcelUtils.writeDataIntoCell(summarySheet, 29, 3, summaryCompanyRDData.RDSpendSourceOfCompany);
            //2.来自政府部门（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 30, 3, summaryCompanyRDData.RDSpendSourceOfGovernment);
            //3.来自银行贷款（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 31, 3, summaryCompanyRDData.RDSpendSourceOfBank);
            //4.来自风险投资（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 32, 3, summaryCompanyRDData.RDSpendSourceOfRiskCapital);
            //5.来自其他渠道（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 33, 3, summaryCompanyRDData.RDSpendSourceOfOthers);

            //五、相关政策落实情况
            //申报加计扣除减免税的研究开发支出(千元)
            ExcelUtils.writeDataIntoCell(summarySheet, 35, 3, summaryCompanyRDData.PolicyImplementDeclareAddtionRD);
            //加计扣除减免税金额(千元)
            ExcelUtils.writeDataIntoCell(summarySheet, 36, 3, summaryCompanyRDData.PolicyImplementAddtionRDTaxFree);
            //高新技术企业减免税金额(千元)
            ExcelUtils.writeDataIntoCell(summarySheet, 37, 3, summaryCompanyRDData.PolicyImplementHighTechRDTaxFree);

            //六、企业办研究开发机构（境内）情况
            //期末机构数(个)
            ExcelUtils.writeDataIntoCell(summarySheet, 5, 10, summaryCompanyRDData.CompanyRunOrgCountEndOfPeriod);
            //机构研究开发人员（人）
            ExcelUtils.writeDataIntoCell(summarySheet, 6, 10, summaryCompanyRDData.CompanyRunOrgRDPersonnel);
            //其中：博士毕业（人）
            ExcelUtils.writeDataIntoCell(summarySheet, 7, 10, summaryCompanyRDData.CompanyRunOrgRDDoctor);
            //其中：硕士毕业（人）
            ExcelUtils.writeDataIntoCell(summarySheet, 8, 10, summaryCompanyRDData.CompanyRunOrgRDMaster);
            //机构研究开发费用（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 9, 10, summaryCompanyRDData.CompanyRunOrgRDExpenses);
            //期末仪器和设备原价（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 10, 10, summaryCompanyRDData.CompanyRunOrgEquipmentValueEndOfPeriod);

            //七、研究开发产出及相关情况
            //(一) 专利情况
            //当年专利申请数(件)
            ExcelUtils.writeDataIntoCell(summarySheet, 13, 10, summaryCompanyRDData.PatentApplyOfCurrentYear);
            //其中：发明专利（件）
            ExcelUtils.writeDataIntoCell(summarySheet, 14, 10, summaryCompanyRDData.PatentApplyOfInvention);
            //期末有效发明专利数（件）
            ExcelUtils.writeDataIntoCell(summarySheet, 15, 10, summaryCompanyRDData.PatentApplyOfInForcePeriod);
            //其中：已被实施（件）
            ExcelUtils.writeDataIntoCell(summarySheet, 16, 10, summaryCompanyRDData.PatentApplyOfBeenImplement);
            //专利所有权转让及许可数（件）
            ExcelUtils.writeDataIntoCell(summarySheet, 17, 10, summaryCompanyRDData.PatentApplyOfAssignment);
            //专利所有权转让及许可收入（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 18, 10, summaryCompanyRDData.PatentApplyOfAssignmentIncome);

            //(二) 新产品情况
            //*新产品销售收入(千元)
            ExcelUtils.writeDataIntoCell(summarySheet, 20, 10, summaryCompanyRDData.NewProductSaleRevenue);
            //*其中：出口(千元)
            ExcelUtils.writeDataIntoCell(summarySheet, 21, 10, summaryCompanyRDData.NewProductSaleOfOutlet);

            //(三)其他情况
            //*期末拥有注册商标(件)
            ExcelUtils.writeDataIntoCell(summarySheet, 23, 10, summaryCompanyRDData.TrademarkOfPeriod);
            //发表科技论文(篇)
            ExcelUtils.writeDataIntoCell(summarySheet, 24, 10, summaryCompanyRDData.ScientificPapers);
            //形成国家或行业标准(项)
            ExcelUtils.writeDataIntoCell(summarySheet, 25, 10, summaryCompanyRDData.StandardsOfNational);

            //八、其他相关情况
            //(一)技术改造和技术获取情况
            //技术改造经费支出（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 28, 10, summaryCompanyRDData.TechTransformExpenses);
            //购买境内技术经费支出（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 29, 10, summaryCompanyRDData.BuyDomesticTechExpenses);
            //引进境外技术经费支出（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 30, 10, summaryCompanyRDData.ImpOverseasTechExpenses);
            //引进境外技术的消化吸收经费支出（千元）
            ExcelUtils.writeDataIntoCell(summarySheet, 31, 10, summaryCompanyRDData.ImpOverseasTechDigestionExpenses);

            // (二)企业办研究开发机构（境外）情况
            //期末企业在境外设立的研究开发机构数(个)
            ExcelUtils.writeDataIntoCell(summarySheet, 33, 10, summaryCompanyRDData.OverseasOrgCount);

            summaryFs.Close();

            //创建汇总文件夹

            using (FileStream output = new FileStream(summaryFilePath + "\\附件1 企业研发活动情况表-汇总.xlsx", FileMode.Create, FileAccess.Write))
            {
                summaryWorkbook.Write(output);
            }
        }

        //读取Excel中的内容并进行数据加减
        public void dealExcelAndPrintLogThreadMethod(DirectoryInfo directoryInfo)
        {
            this.BeginInvoke(dealExcelAndPrintLogDelegate, directoryInfo);
        }
    }
}
