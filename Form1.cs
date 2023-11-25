using Sunny.UI;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Formula.Functions;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Sunny.UI.Win32;
using NPOI.Util;

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

        }

        //选择文件夹目录
        private void chooseFolderBtn_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            folderBrowserDialog.ShowDialog();
            this.folderPathText.Text = folderBrowserDialog.SelectedPath;
        }

        ////获取指定目录下的Excel，然后汇总对应的数据
        private void summaryBtn_Click(object sender, EventArgs e)
        {
            //清空日志，以便打印新的
            this.logTextBox.Clear();
            this.errorLogTextBox.Clear();

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
                FileOptUtils.DeleteAllFiles(errorFilePath);
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

                //是否要
                Boolean isSummary = true;
                //错误日志信息
                String errorText = "";
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
                        errorLogTextBox.AppendText("不支持的文件：" + fsInfo.Name + "\r\n");
                        continue;
                    }

                    //获取Excel中的数据
                    // 第一个工作表对象（107-1：企业研发项目填报报表）
                    ISheet RDProjectSheet = workbook.GetSheetAt(0);
                    // 第二个工作表对象（107-2：企业研发活动相关情况表）
                    ISheet RDAcitivitySheet = workbook.GetSheetAt(1);
                    //开始获取第一个表格（107-1：企业研发项目填报报表）的数据
                    Get1071SheetData(companyRDData, RDProjectSheet);
                    //开始获取第二个Sheet表（107-2：企业研发活动相关情况表）的数据
                    Get1072SheetData(companyRDData, RDAcitivitySheet);

                    //逻辑判断
                    LogicCheck(companyRDData, ref isSummary, ref errorText);

                    //如果逻辑校验没有问题
                    if (isSummary)
                    {
                        companyRDDatas.Add(companyRDData);
                        //提醒不剔除数据的
                        if (!string.IsNullOrWhiteSpace(errorText)) { 
                            errorLogTextBox.AppendText("常规提示：《" + fsInfo.Name+"》" +"中触发了以下规则："+errorText+ "\r\n");
                        }
                    }
                    else { 
                        errorLogTextBox.AppendText("必须修改：《" + fsInfo.Name+"》" +"中触发了以下规则："+errorText+ "\r\n");
                        //拷贝一份文件到异常数据文件夹
                        File.Copy(fsInfo.FullName, errorFilePath + fsInfo.Name);
                    }

                    fs.Close();
                }
                else { 
                    errorLogTextBox.AppendText("不支持的文件：" + fsInfo.Name + "\r\n");
                }
            }

            //开始遍历并合并数据
            CompanyRDData summaryCompanyRDData = new CompanyRDData();
            summaryDataFun(companyRDDatas, summaryCompanyRDData);

            //导出数据到Excel中
            exportSummaryDataIntoExcel(summaryCompanyRDData, summaryFilePath);
        }

        //获取107-1表中的数据并赋值到对象的列表中
        private void Get1071SheetData(CompanyRDData companyRDData, ISheet RDProjectSheet) {

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
        private void LogicCheck(CompanyRDData companyRDData, ref Boolean isSummary, ref String errorText) {
            //下面这些是将数据剔除出去的条件
            //1≥2（研究开发人员合计≥其中：管理和服务人员）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelManageAndService)
            {
                isSummary = false;
                errorText += "研究开发人员合计≥其中：管理和服务人员；";
            }
            //1≥3（研究开发人员合计≥其中：女性）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelFemale)
            {
                isSummary = false;
                errorText += "研究开发人员合计≥其中：女性；";
            }
            //1≥4（研究开发人员合计≥其中：全职人员）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelFullTimeStaff)
            {
                isSummary = false;
                errorText += "研究开发人员合计≥其中：全职人员；";
            }
            //1≥5≥24+25（研究开发人员合计≥其中：本科毕业及以上人员≥其中：博士毕业+其中：硕士毕业）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelBachelorAndAbove || companyRDData.RDPersonnelBachelorAndAbove < (companyRDData.CompanyRunOrgRDDoctor + companyRDData.CompanyRunOrgRDMaster))
            {
                isSummary = false;
                errorText += "研究开发人员合计≥其中：本科毕业及以上人员≥其中：博士毕业+其中：硕士毕业；";
            }
            //1≥6（研究开发人员合计≥其中：外聘人员）
            if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelExternalStaff)
            {
                isSummary = false;
                errorText += "研究开发人员合计≥其中：外聘人员；";
            }
            //1≥23≥24+25（研究开发人员合计≥机构研究开发人员≥其中：博士毕业+其中：硕士毕业）
            if (companyRDData.RDPersonnelTotal < companyRDData.CompanyRunOrgRDPersonnel || companyRDData.CompanyRunOrgRDPersonnel < (companyRDData.CompanyRunOrgRDDoctor + companyRDData.CompanyRunOrgRDMaster))
            {
                isSummary = false;
                errorText += "研究开发人员合计≥机构研究开发人员≥其中：博士毕业+其中：硕士毕业；";
            }
            //7=8+9+10+11+12+13+14+19≥26（研究开发费用合计=人员人工费用+直接投入费用+折旧费用与长期待摊费用+无形资产摊销费用+设计费用+装备调试费用与试验费用+委托外部研究开发费用+其他费用≥机构研究开发费用）
            if (companyRDData.RDExpensesTotal != (companyRDData.RDExpensesPersonnelLabor + companyRDData.RDExpensesDirectInput + companyRDData.RDExpensesDepreciationAndLongTerm + companyRDData.RDExpensesIntangibleAssets + companyRDData.RDExpensesDesign + companyRDData.RDExpensesEquipmentDebug + companyRDData.RDExpensesEntrustOutsourcedRD + companyRDData.RDExpensesOthers) || companyRDData.RDExpensesTotal < companyRDData.CompanyRunOrgRDExpenses)
            {
                isSummary = false;
                errorText += "研究开发费用合计=人员人工费用+直接投入费用+折旧费用与长期待摊费用+无形资产摊销费用+设计费用+装备调试费用与试验费用+委托外部研究开发费用+其他费用≥机构研究开发费用；";
            }
            //若1>0，则8>0（若研究开发人员合计>0，则人员人工费用>0）
            if (companyRDData.RDPersonnelTotal > 0 && companyRDData.RDExpensesPersonnelLabor == 0)
            {
                isSummary = false;
                errorText += "若研究开发人员合计>0，则人员人工费用>0；";
            }
            //若8>0，则1>0（若人员人工费用>0，则研究开发人员合计>0）
            if (companyRDData.RDExpensesPersonnelLabor > 0 && companyRDData.RDPersonnelTotal == 0)
            {
                isSummary = false;
                errorText += "若人员人工费用>0，则研究开发人员合计>0；";
            }
            //14=15+16+17+18（委托外部研究开发费用=①委托境内研究机构+②委托境内高等学校+③委托境内企业+④委托境外机构）
            if (companyRDData.RDExpensesEntrustOutsourcedRD != (companyRDData.RDExpensesEntrustDomesticResearch + companyRDData.RDExpensesEntrustDomesticCollege + companyRDData.RDExpensesEntrustDomesticCompany + companyRDData.RDExpensesEntrustOverseasInstitutions))
            {
                isSummary = false;
                errorText += "委托外部研究开发费用=①委托境内研究机构+②委托境内高等学校+③委托境内企业+④委托境外机构；";
            }
            //20≥21（当年形成用于研究开发的固定资产≥其中：仪器和设备）
            if (companyRDData.RDAssetsYear < companyRDData.RDAssetsYearEquipment)
            {
                isSummary = false;
                errorText += "当年形成用于研究开发的固定资产≥其中：仪器和设备；";
            }
            //若27>0，则22>0（若期末仪器和设备原价>0，则期末机构数>0）
            if (companyRDData.CompanyRunOrgEquipmentValueEndOfPeriod > 0 && companyRDData.CompanyRunOrgCountEndOfPeriod == 0)
            {
                isSummary = false;
                errorText += "若期末仪器和设备原价>0，则期末机构数>0；";
            }
            //29≥30（当年专利申请数≥其中：发明专利）
            if (companyRDData.PatentApplyOfCurrentYear < companyRDData.PatentApplyOfInvention)
            {
                isSummary = false;
                errorText += "当年专利申请数≥其中：发明专利；";
            }
            //32≥33（期末有效发明专利数≥其中：已被实施）
            if (companyRDData.PatentApplyOfInForcePeriod < companyRDData.PatentApplyOfBeenImplement)
            {
                isSummary = false;
                errorText += "期末有效发明专利数≥其中：已被实施；";
            }
            //36≥37（新产品销售收入≥其中：出口）
            if (companyRDData.NewProductSaleRevenue < companyRDData.NewProductSaleOfOutlet)
            {
                isSummary = false;
                errorText += "新产品销售收入≥其中：出口；";
            }
            //研究开发费用合计 = 四、研究开发支出资金来源中各项的和
            if (companyRDData.RDExpensesTotal != (companyRDData.RDSpendSourceOfCompany + companyRDData.RDSpendSourceOfGovernment + companyRDData.RDSpendSourceOfBank + companyRDData.RDSpendSourceOfRiskCapital + companyRDData.RDSpendSourceOfOthers))
            {
                isSummary = false;
                errorText += "研究开发费用合计 = 四、研究开发支出资金来源中各项的和；";
            }

            //下面这些是只做提醒的的条件
            //"期末机构数"如果为0，则进行提醒，不剔除数据
            if (companyRDData.CompanyRunOrgCountEndOfPeriod == 0)
            {
                errorText += "期末机构数为0；";
            }
            //" 7.委托外部研究开发费用"如果大于0，，则进行提醒，不剔除数据
            if (companyRDData.RDExpensesEntrustOutsourcedRD > 0)
            {
                errorText += "存在委托外部研究开发费用；";
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
