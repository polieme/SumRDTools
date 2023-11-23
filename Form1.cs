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
                    Thread myThread = new Thread(new ThreadStart(async delegate
                    {
                        dealExcelAndPrintLogThreadMethod(directoryInfo);
                    }));
                    myThread.Start();
                }
            }
        }

        //读取Excel中内容并打印日志的方法（传过来的参数是选择目录下的信息对象）
        private void dealExcelAndPrintLogFun(DirectoryInfo directoryInfo)
        {
            //创建汇总文件夹和错误文件文件夹
            string summaryFilePath = directoryInfo.FullName + "\\汇总";
            string errorFilePath = directoryInfo.FullName + "\\异常数据\\";
            if (!Directory.Exists(summaryFilePath))
            {
                Directory.CreateDirectory(summaryFilePath);
            }
            else {
                DeleteAllFiles(summaryFilePath);
            }
            if (!Directory.Exists(errorFilePath))
            {
                Directory.CreateDirectory(errorFilePath);
            }
            else { 
                DeleteAllFiles(errorFilePath);
            }


            FileSystemInfo[] fsInfos = directoryInfo.GetFiles();
            //定义对象存储Excel中的数据
            List< CompanyRDData > companyRDDatas = new List< CompanyRDData >();
            foreach (FileSystemInfo fsInfo in fsInfos)
            {
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
                    FileStream fs = new FileStream(fsInfo.FullName, FileMode.Open, FileAccess.Read);
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

                    // 获取第一个工作表
                    ISheet sheet = workbook.GetSheetAt(1);

                    //研究开发人员合计(人)
                    companyRDData.RDPersonnelTotal = getInt(getCellValueByCellType(sheet, 5, 3));
                    //其中：管理和服务人员
                    companyRDData.RDPersonnelManageAndService = getInt(getCellValueByCellType(sheet, 6, 3));
                    //其中：女性(人)
                    companyRDData.RDPersonnelFemale = getInt(getCellValueByCellType(sheet, 7, 3));
                    //其中：全职人员
                    companyRDData.RDPersonnelFullTimeStaff = getInt(getCellValueByCellType(sheet, 8, 3));
                    //其中：本科毕业及以上人员(人)
                    companyRDData.RDPersonnelBachelorAndAbove = getInt(getCellValueByCellType(sheet, 9, 3));
                    //其中：外聘人员(人)
                    companyRDData.RDPersonnelExternalStaff = getInt(getCellValueByCellType(sheet, 10, 3));

                    //二、研究开发费用情况
                    // 研究开发费用合计（千元）
                    companyRDData.RDExpensesTotal = getFloat(getCellValueByCellType(sheet, 12, 3));
                    //1.人员人工费用（千元）
                    companyRDData.RDExpensesPersonnelLabor = getFloat(getCellValueByCellType(sheet, 13, 3));
                    //2.直接投入费用（千元）
                    companyRDData.RDExpensesDirectInput = getFloat(getCellValueByCellType(sheet, 14, 3));
                    //3.折旧费用与长期待摊费用（千元）
                    companyRDData.RDExpensesDepreciationAndLongTerm = getFloat(getCellValueByCellType(sheet, 15, 3));
                    //4.无形资产摊销费用（千元）
                    companyRDData.RDExpensesIntangibleAssets = getFloat(getCellValueByCellType(sheet, 16, 3));
                    //5.设计费用（千元）
                    companyRDData.RDExpensesDesign = getFloat(getCellValueByCellType(sheet, 17, 3));
                    //6.装备调试费用与试验费用（千元）
                    companyRDData.RDExpensesEquipmentDebug = getFloat(getCellValueByCellType(sheet, 18, 3));
                    //7.委托外部研究开发费用（千元）
                    companyRDData.RDExpensesEntrustOutsourcedRD = getFloat(getCellValueByCellType(sheet, 19, 3));
                    //①委托境内研究机构（千元）
                    companyRDData.RDExpensesEntrustDomesticResearch = getFloat(getCellValueByCellType(sheet, 20, 3));
                    //②委托境内高等学校（千元）
                    companyRDData.RDExpensesEntrustDomesticCollege = getFloat(getCellValueByCellType(sheet, 21, 3));
                    //③委托境内企业（千元）
                    companyRDData.RDExpensesEntrustDomesticCompany = getFloat(getCellValueByCellType(sheet, 22, 3));
                    //④委托境外机构（千元）
                    companyRDData.RDExpensesEntrustOverseasInstitutions = getFloat(getCellValueByCellType(sheet, 23, 3));
                    //8.其他费用（千元）
                    companyRDData.RDExpensesOthers = getFloat(getCellValueByCellType(sheet, 24, 3));

                    //三、研究开发资产情况
                    //当年形成用于研究开发的固定资产（千元）
                    companyRDData.RDAssetsYear = getFloat(getCellValueByCellType(sheet, 26, 3));
                    //其中：仪器和设备（千元）
                    companyRDData.RDAssetsYearEquipment = getFloat(getCellValueByCellType(sheet, 27, 3));

                    //四、研究开发支出资金来源
                    //1.来自企业自筹(千元)
                    companyRDData.RDSpendSourceOfCompany = getFloat(getCellValueByCellType(sheet, 29, 3));
                    //2.来自政府部门（千元）
                    companyRDData.RDSpendSourceOfGovernment = getFloat(getCellValueByCellType(sheet, 30, 3));
                    //3.来自银行贷款（千元）
                    companyRDData.RDSpendSourceOfBank = getFloat(getCellValueByCellType(sheet, 31, 3));
                    //4.来自风险投资（千元）
                    companyRDData.RDSpendSourceOfRiskCapital = getFloat(getCellValueByCellType(sheet, 32, 3));
                    //5.来自其他渠道（千元）
                    companyRDData.RDSpendSourceOfOthers = getFloat(getCellValueByCellType(sheet, 33, 3));

                    //五、相关政策落实情况
                    //申报加计扣除减免税的研究开发支出(千元)
                    companyRDData.PolicyImplementDeclareAddtionRD = getFloat(getCellValueByCellType(sheet, 35, 3));
                    //加计扣除减免税金额(千元)
                    companyRDData.PolicyImplementAddtionRDTaxFree = getFloat(getCellValueByCellType(sheet, 36, 3));
                    //高新技术企业减免税金额(千元)
                    companyRDData.PolicyImplementHighTechRDTaxFree = getFloat(getCellValueByCellType(sheet, 37, 3));

                    //六、企业办研究开发机构（境内）情况
                    //期末机构数(个)
                    companyRDData.CompanyRunOrgCountEndOfPeriod = getInt(getCellValueByCellType(sheet, 5, 10));
                    //机构研究开发人员（人）
                    companyRDData.CompanyRunOrgRDPersonnel = getInt(getCellValueByCellType(sheet, 6, 10));
                    //其中：博士毕业（人）
                    companyRDData.CompanyRunOrgRDDoctor = getInt(getCellValueByCellType(sheet, 7, 10));
                    //其中：硕士毕业（人）
                    companyRDData.CompanyRunOrgRDMaster = getInt(getCellValueByCellType(sheet, 8, 10));
                    //机构研究开发费用（千元）
                    companyRDData.CompanyRunOrgRDExpenses = getFloat(getCellValueByCellType(sheet, 9, 10));
                    //期末仪器和设备原价（千元）
                    companyRDData.CompanyRunOrgEquipmentValueEndOfPeriod = getFloat(getCellValueByCellType(sheet, 10, 10));

                    //七、研究开发产出及相关情况
                    //(一) 专利情况
                    //当年专利申请数(件)
                    companyRDData.PatentApplyOfCurrentYear = getInt(getCellValueByCellType(sheet, 13, 10));
                    //其中：发明专利（件）
                    companyRDData.PatentApplyOfInvention = getInt(getCellValueByCellType(sheet, 14, 10));
                    //期末有效发明专利数（件）
                    companyRDData.PatentApplyOfInForcePeriod = getInt(getCellValueByCellType(sheet, 15, 10));
                    //其中：已被实施（件）
                    companyRDData.PatentApplyOfBeenImplement = getInt(getCellValueByCellType(sheet, 16, 10));
                    //专利所有权转让及许可数（件）
                    companyRDData.PatentApplyOfAssignment = getInt(getCellValueByCellType(sheet, 17, 10));
                    //专利所有权转让及许可收入（千元）
                    companyRDData.PatentApplyOfAssignmentIncome = getFloat(getCellValueByCellType(sheet, 18, 10));
                    //(二) 新产品情况
                    //*新产品销售收入(千元)
                    companyRDData.NewProductSaleRevenue = getFloat(getCellValueByCellType(sheet, 20, 10));
                    //*其中：出口(千元)
                    companyRDData.NewProductSaleOfOutlet = getFloat(getCellValueByCellType(sheet, 21, 10));
                    //(三)其他情况
                    //*期末拥有注册商标(件)
                    companyRDData.TrademarkOfPeriod = getInt(getCellValueByCellType(sheet, 23, 10));
                    //发表科技论文(篇)
                    companyRDData.ScientificPapers = getInt(getCellValueByCellType(sheet, 24, 10));
                    //形成国家或行业标准(项)
                    companyRDData.StandardsOfNational = getInt(getCellValueByCellType(sheet, 25, 10));

                    //八、其他相关情况
                    //(一)技术改造和技术获取情况
                    //技术改造经费支出（千元）
                    companyRDData.TechTransformExpenses = getFloat(getCellValueByCellType(sheet, 28, 10));
                    //购买境内技术经费支出（千元）
                    companyRDData.BuyDomesticTechExpenses = getFloat(getCellValueByCellType(sheet, 29, 10));
                    //引进境外技术经费支出（千元）
                    companyRDData.ImpOverseasTechExpenses = getFloat(getCellValueByCellType(sheet, 30, 10));
                    //引进境外技术的消化吸收经费支出（千元）
                    companyRDData.ImpOverseasTechDigestionExpenses = getFloat(getCellValueByCellType(sheet, 31, 10));
                    // (二)企业办研究开发机构（境外）情况
                    //期末企业在境外设立的研究开发机构数(个)
                    companyRDData.OverseasOrgCount = getInt(getCellValueByCellType(sheet, 33, 10));


                    //逻辑判断
                    //1≥2（研究开发人员合计≥其中：管理和服务人员）
                    if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelManageAndService) {
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
                    if (companyRDData.RDPersonnelTotal < companyRDData.RDPersonnelBachelorAndAbove || companyRDData.RDPersonnelBachelorAndAbove < (companyRDData.CompanyRunOrgRDDoctor+ companyRDData.CompanyRunOrgRDMaster))
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
                    if (companyRDData.RDPersonnelTotal < companyRDData.CompanyRunOrgRDPersonnel || companyRDData.CompanyRunOrgRDPersonnel<(companyRDData.CompanyRunOrgRDDoctor+ companyRDData.CompanyRunOrgRDMaster))
                    {
                        isSummary = false;
                        errorText += "研究开发人员合计≥机构研究开发人员≥其中：博士毕业+其中：硕士毕业；";
                    }
                    //7=8+9+10+11+12+13+14+19≥26（研究开发费用合计=人员人工费用+直接投入费用+折旧费用与长期待摊费用+无形资产摊销费用+设计费用+装备调试费用与试验费用+委托外部研究开发费用+其他费用≥机构研究开发费用）
                    if (new Decimal(companyRDData.RDExpensesTotal) != (new Decimal(companyRDData.RDExpensesPersonnelLabor) + new Decimal(companyRDData.RDExpensesDirectInput) + new Decimal(companyRDData.RDExpensesDepreciationAndLongTerm) + new Decimal(companyRDData.RDExpensesIntangibleAssets) + new Decimal(companyRDData.RDExpensesDesign) + new Decimal(companyRDData.RDExpensesEquipmentDebug) + new Decimal(companyRDData.RDExpensesEntrustOutsourcedRD) + new Decimal(companyRDData.RDExpensesOthers)) || companyRDData.RDExpensesTotal < companyRDData.CompanyRunOrgRDExpenses)
                    {
                        isSummary = false;
                        errorText += "研究开发费用合计=人员人工费用+直接投入费用+折旧费用与长期待摊费用+无形资产摊销费用+设计费用+装备调试费用与试验费用+委托外部研究开发费用+其他费用≥机构研究开发费用；";
                    }
                    //若1>0，则8>0（若研究开发人员合计>0，则人员人工费用>0）
                    if (companyRDData.RDPersonnelTotal > 0 && companyRDData.RDExpensesPersonnelLabor == 0) {
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
                    if (new Decimal(companyRDData.RDExpensesEntrustOutsourcedRD) != (new Decimal(companyRDData.RDExpensesEntrustDomesticResearch) + new Decimal(companyRDData.RDExpensesEntrustDomesticCollege) + new Decimal(companyRDData.RDExpensesEntrustDomesticCompany) + new Decimal(companyRDData.RDExpensesEntrustOverseasInstitutions))) {
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

                    //如果逻辑校验没有问题
                    if (isSummary)
                    {
                        companyRDDatas.Add(companyRDData);
                    }
                    else { 
                        errorLogTextBox.AppendText("《"+fsInfo.Name+"》" +"中触发了以下规则："+errorText+ "\r\n");
                        //拷贝一份文件到触发规则文件夹
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
            foreach (var companyRDData in companyRDDatas) {
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
                summaryCompanyRDData.RDExpensesTotal = (float)(new Decimal(summaryCompanyRDData.RDExpensesTotal) + new Decimal(companyRDData.RDExpensesTotal));
                //1.人员人工费用（千元）
                summaryCompanyRDData.RDExpensesPersonnelLabor = (float)(new Decimal(summaryCompanyRDData.RDExpensesPersonnelLabor) + new Decimal(companyRDData.RDExpensesPersonnelLabor));
                //2.直接投入费用（千元）
                summaryCompanyRDData.RDExpensesDirectInput = (float)(new Decimal(summaryCompanyRDData.RDExpensesDirectInput) + new Decimal(companyRDData.RDExpensesDirectInput));
                //3.折旧费用与长期待摊费用（千元）
                summaryCompanyRDData.RDExpensesDepreciationAndLongTerm = (float)(new Decimal(summaryCompanyRDData.RDExpensesDepreciationAndLongTerm) + new Decimal(companyRDData.RDExpensesDepreciationAndLongTerm));
                //4.无形资产摊销费用（千元）
                summaryCompanyRDData.RDExpensesIntangibleAssets = (float)(new Decimal(summaryCompanyRDData.RDExpensesIntangibleAssets) + new Decimal(companyRDData.RDExpensesIntangibleAssets));

                //5.设计费用（千元）
                summaryCompanyRDData.RDExpensesDesign = (float)(new Decimal(summaryCompanyRDData.RDExpensesDesign) + new Decimal(companyRDData.RDExpensesDesign));

                //6.装备调试费用与试验费用（千元）
                summaryCompanyRDData.RDExpensesEquipmentDebug = (float)(new Decimal(summaryCompanyRDData.RDExpensesEquipmentDebug) + new Decimal(companyRDData.RDExpensesEquipmentDebug));
                //7.委托外部研究开发费用（千元）
                summaryCompanyRDData.RDExpensesEntrustOutsourcedRD = (float)(new Decimal(summaryCompanyRDData.RDExpensesEntrustOutsourcedRD) + new Decimal(companyRDData.RDExpensesEntrustOutsourcedRD));
                //①委托境内研究机构（千元）
                summaryCompanyRDData.RDExpensesEntrustDomesticResearch = (float)(new Decimal(summaryCompanyRDData.RDExpensesEntrustDomesticResearch) + new Decimal(companyRDData.RDExpensesEntrustDomesticResearch));
                //②委托境内高等学校（千元）
                summaryCompanyRDData.RDExpensesEntrustDomesticCollege = (float)(new Decimal(summaryCompanyRDData.RDExpensesEntrustDomesticCollege) + new Decimal(companyRDData.RDExpensesEntrustDomesticCollege));
                //③委托境内企业（千元）
                summaryCompanyRDData.RDExpensesEntrustDomesticCompany = (float)(new Decimal(summaryCompanyRDData.RDExpensesEntrustDomesticCompany) + new Decimal(companyRDData.RDExpensesEntrustDomesticCompany));
                //④委托境外机构（千元）
                summaryCompanyRDData.RDExpensesEntrustOverseasInstitutions = (float)(new Decimal(summaryCompanyRDData.RDExpensesEntrustOverseasInstitutions) + new Decimal(companyRDData.RDExpensesEntrustOverseasInstitutions));
                //8.其他费用（千元）
                summaryCompanyRDData.RDExpensesOthers = (float)(new Decimal(summaryCompanyRDData.RDExpensesOthers) + new Decimal(companyRDData.RDExpensesOthers));

                //三、研究开发资产情况
                //当年形成用于研究开发的固定资产（千元）
                summaryCompanyRDData.RDAssetsYear = (float)(new Decimal(summaryCompanyRDData.RDAssetsYear) + new Decimal(companyRDData.RDAssetsYear));
                //其中：仪器和设备（千元）
                summaryCompanyRDData.RDAssetsYearEquipment = (float)(new Decimal(summaryCompanyRDData.RDAssetsYearEquipment) + new Decimal(companyRDData.RDAssetsYearEquipment));


                //四、研究开发支出资金来源
                //1.来自企业自筹(千元)
                summaryCompanyRDData.RDSpendSourceOfCompany = (float)(new Decimal(summaryCompanyRDData.RDSpendSourceOfCompany) + new Decimal(companyRDData.RDSpendSourceOfCompany));
                //2.来自政府部门（千元）
                summaryCompanyRDData.RDSpendSourceOfGovernment = (float)(new Decimal(summaryCompanyRDData.RDSpendSourceOfGovernment) + new Decimal(companyRDData.RDSpendSourceOfGovernment));
                //3.来自银行贷款（千元）
                summaryCompanyRDData.RDSpendSourceOfBank = (float)(new Decimal(summaryCompanyRDData.RDSpendSourceOfBank) + new Decimal(companyRDData.RDSpendSourceOfBank));
                //4.来自风险投资（千元）
                summaryCompanyRDData.RDSpendSourceOfRiskCapital = (float)(new Decimal(summaryCompanyRDData.RDSpendSourceOfRiskCapital) + new Decimal(companyRDData.RDSpendSourceOfRiskCapital));
                //5.来自其他渠道（千元）
                summaryCompanyRDData.RDSpendSourceOfOthers = (float)(new Decimal(summaryCompanyRDData.RDSpendSourceOfOthers) + new Decimal(companyRDData.RDSpendSourceOfOthers));

                //五、相关政策落实情况
                //申报加计扣除减免税的研究开发支出(千元)
                summaryCompanyRDData.PolicyImplementDeclareAddtionRD = (float)(new Decimal(summaryCompanyRDData.PolicyImplementDeclareAddtionRD) + new Decimal(companyRDData.PolicyImplementDeclareAddtionRD));
                //加计扣除减免税金额(千元)
                summaryCompanyRDData.PolicyImplementAddtionRDTaxFree = (float)(new Decimal(summaryCompanyRDData.PolicyImplementAddtionRDTaxFree) + new Decimal(companyRDData.PolicyImplementAddtionRDTaxFree));
                //高新技术企业减免税金额(千元)
                summaryCompanyRDData.PolicyImplementHighTechRDTaxFree = (float)(new Decimal(summaryCompanyRDData.PolicyImplementHighTechRDTaxFree) + new Decimal(companyRDData.PolicyImplementHighTechRDTaxFree));


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
                summaryCompanyRDData.CompanyRunOrgRDExpenses = (float)(new Decimal(summaryCompanyRDData.CompanyRunOrgRDExpenses) + new Decimal(companyRDData.CompanyRunOrgRDExpenses));
                //期末仪器和设备原价（千元）
                summaryCompanyRDData.CompanyRunOrgEquipmentValueEndOfPeriod = (float)(new Decimal(summaryCompanyRDData.CompanyRunOrgEquipmentValueEndOfPeriod) + new Decimal(companyRDData.CompanyRunOrgEquipmentValueEndOfPeriod));

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
                summaryCompanyRDData.PatentApplyOfAssignmentIncome = (float)(new Decimal(summaryCompanyRDData.PatentApplyOfAssignmentIncome) + new Decimal(companyRDData.PatentApplyOfAssignmentIncome));

                //(二) 新产品情况
                //*新产品销售收入(千元)
                summaryCompanyRDData.NewProductSaleRevenue = (float)(new Decimal(summaryCompanyRDData.NewProductSaleRevenue) + new Decimal(companyRDData.NewProductSaleRevenue));
                //*其中：出口(千元)
                summaryCompanyRDData.NewProductSaleOfOutlet = (float)(new Decimal(summaryCompanyRDData.NewProductSaleOfOutlet) + new Decimal(companyRDData.NewProductSaleOfOutlet));
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
                summaryCompanyRDData.TechTransformExpenses = (float)(new Decimal(summaryCompanyRDData.TechTransformExpenses) + new Decimal(companyRDData.TechTransformExpenses));
                //购买境内技术经费支出（千元）
                summaryCompanyRDData.BuyDomesticTechExpenses = (float)(new Decimal(summaryCompanyRDData.BuyDomesticTechExpenses) + new Decimal(companyRDData.BuyDomesticTechExpenses));
                //引进境外技术经费支出（千元）
                summaryCompanyRDData.ImpOverseasTechExpenses = (float)(new Decimal(summaryCompanyRDData.ImpOverseasTechExpenses) + new Decimal(companyRDData.ImpOverseasTechExpenses));
                //引进境外技术的消化吸收经费支出（千元）
                summaryCompanyRDData.ImpOverseasTechDigestionExpenses = (float)(new Decimal(summaryCompanyRDData.ImpOverseasTechDigestionExpenses) + new Decimal(companyRDData.ImpOverseasTechDigestionExpenses));
                // (二)企业办研究开发机构（境外）情况
                //期末企业在境外设立的研究开发机构数(个)
                summaryCompanyRDData.OverseasOrgCount += companyRDData.OverseasOrgCount;
            }

            //导出Excel
            FileStream summaryFs = new FileStream("EnterprisesRDTemplate.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook summaryWorkbook = new XSSFWorkbook(summaryFs);
            ISheet summarySheet = summaryWorkbook.GetSheetAt(1);
            //填写数据到导出的Excel中
            //研究开发人员合计(人)
            writeDataIntoCell(summarySheet, 5, 3, summaryCompanyRDData.RDPersonnelTotal);
            //其中：管理和服务人员
            writeDataIntoCell(summarySheet, 6, 3, summaryCompanyRDData.RDPersonnelManageAndService);
            //其中：女性(人)
            writeDataIntoCell(summarySheet, 7, 3, summaryCompanyRDData.RDPersonnelFemale);
            //其中：全职人员
            writeDataIntoCell(summarySheet, 8, 3, summaryCompanyRDData.RDPersonnelFullTimeStaff);
            //其中：本科毕业及以上人员(人)
            writeDataIntoCell(summarySheet, 9, 3, summaryCompanyRDData.RDPersonnelBachelorAndAbove);
            //其中：外聘人员(人)
            writeDataIntoCell(summarySheet, 10, 3, summaryCompanyRDData.RDPersonnelExternalStaff);


            //二、研究开发费用情况
            // 研究开发费用合计（千元）
            writeDataIntoCell(summarySheet, 12, 3, summaryCompanyRDData.RDExpensesTotal);
            //1.人员人工费用（千元）
            writeDataIntoCell(summarySheet, 13, 3, summaryCompanyRDData.RDExpensesPersonnelLabor);
            //2.直接投入费用（千元）
            writeDataIntoCell(summarySheet, 14, 3, summaryCompanyRDData.RDExpensesDirectInput);
            //3.折旧费用与长期待摊费用（千元）
            writeDataIntoCell(summarySheet, 15, 3, summaryCompanyRDData.RDExpensesDepreciationAndLongTerm);
            //4.无形资产摊销费用（千元）
            writeDataIntoCell(summarySheet, 16, 3, summaryCompanyRDData.RDExpensesIntangibleAssets);
            //5.设计费用（千元）
            writeDataIntoCell(summarySheet, 17, 3, summaryCompanyRDData.RDExpensesDesign);
            //6.装备调试费用与试验费用（千元）
            writeDataIntoCell(summarySheet, 18, 3, summaryCompanyRDData.RDExpensesEquipmentDebug);
            //7.委托外部研究开发费用（千元）
            writeDataIntoCell(summarySheet, 19, 3, summaryCompanyRDData.RDExpensesEntrustOutsourcedRD);
            //①委托境内研究机构（千元）
            writeDataIntoCell(summarySheet, 20, 3, summaryCompanyRDData.RDExpensesEntrustDomesticResearch);
            //②委托境内高等学校（千元）
            writeDataIntoCell(summarySheet, 21, 3, summaryCompanyRDData.RDExpensesEntrustDomesticCollege);
            //③委托境内企业（千元）
            writeDataIntoCell(summarySheet, 22, 3, summaryCompanyRDData.RDExpensesEntrustDomesticCompany);
            //④委托境外机构（千元）
            writeDataIntoCell(summarySheet, 23, 3, summaryCompanyRDData.RDExpensesEntrustOverseasInstitutions);
            //8.其他费用（千元）
            writeDataIntoCell(summarySheet, 24, 3, summaryCompanyRDData.RDExpensesOthers);


            //三、研究开发资产情况
            //当年形成用于研究开发的固定资产（千元）
            writeDataIntoCell(summarySheet, 26, 3, summaryCompanyRDData.RDAssetsYear);
            //其中：仪器和设备（千元）
            writeDataIntoCell(summarySheet, 27, 3, summaryCompanyRDData.RDAssetsYearEquipment);


            //四、研究开发支出资金来源
            //1.来自企业自筹(千元)
            writeDataIntoCell(summarySheet, 29, 3, summaryCompanyRDData.RDSpendSourceOfCompany);
            //2.来自政府部门（千元）
            writeDataIntoCell(summarySheet, 30, 3, summaryCompanyRDData.RDSpendSourceOfGovernment);
            //3.来自银行贷款（千元）
            writeDataIntoCell(summarySheet, 31, 3, summaryCompanyRDData.RDSpendSourceOfBank);
            //4.来自风险投资（千元）
            writeDataIntoCell(summarySheet, 32, 3, summaryCompanyRDData.RDSpendSourceOfRiskCapital);
            //5.来自其他渠道（千元）
            writeDataIntoCell(summarySheet, 33, 3, summaryCompanyRDData.RDSpendSourceOfOthers);

            //五、相关政策落实情况
            //申报加计扣除减免税的研究开发支出(千元)
            writeDataIntoCell(summarySheet, 35, 3, summaryCompanyRDData.PolicyImplementDeclareAddtionRD);
            //加计扣除减免税金额(千元)
            writeDataIntoCell(summarySheet, 36, 3, summaryCompanyRDData.PolicyImplementAddtionRDTaxFree);
            //高新技术企业减免税金额(千元)
            writeDataIntoCell(summarySheet, 37, 3, summaryCompanyRDData.PolicyImplementHighTechRDTaxFree);

            //六、企业办研究开发机构（境内）情况
            //期末机构数(个)
            writeDataIntoCell(summarySheet, 5, 10, summaryCompanyRDData.CompanyRunOrgCountEndOfPeriod);
            //机构研究开发人员（人）
            writeDataIntoCell(summarySheet, 6, 10, summaryCompanyRDData.CompanyRunOrgRDPersonnel);
            //其中：博士毕业（人）
            writeDataIntoCell(summarySheet, 7, 10, summaryCompanyRDData.CompanyRunOrgRDDoctor);
            //其中：硕士毕业（人）
            writeDataIntoCell(summarySheet, 8, 10, summaryCompanyRDData.CompanyRunOrgRDMaster);
            //机构研究开发费用（千元）
            writeDataIntoCell(summarySheet, 9, 10, summaryCompanyRDData.CompanyRunOrgRDExpenses);
            //期末仪器和设备原价（千元）
            writeDataIntoCell(summarySheet, 10, 10, summaryCompanyRDData.CompanyRunOrgEquipmentValueEndOfPeriod);

            //七、研究开发产出及相关情况
            //(一) 专利情况
            //当年专利申请数(件)
            writeDataIntoCell(summarySheet, 13, 10, summaryCompanyRDData.PatentApplyOfCurrentYear);
            //其中：发明专利（件）
            writeDataIntoCell(summarySheet, 14, 10, summaryCompanyRDData.PatentApplyOfInvention);
            //期末有效发明专利数（件）
            writeDataIntoCell(summarySheet, 15, 10, summaryCompanyRDData.PatentApplyOfInForcePeriod);
            //其中：已被实施（件）
            writeDataIntoCell(summarySheet, 16, 10, summaryCompanyRDData.PatentApplyOfBeenImplement);
            //专利所有权转让及许可数（件）
            writeDataIntoCell(summarySheet, 17, 10, summaryCompanyRDData.PatentApplyOfAssignment);
            //专利所有权转让及许可收入（千元）
            writeDataIntoCell(summarySheet, 18, 10, summaryCompanyRDData.PatentApplyOfAssignmentIncome);
            
            //(二) 新产品情况
            //*新产品销售收入(千元)
            writeDataIntoCell(summarySheet, 20, 10, summaryCompanyRDData.NewProductSaleRevenue);
            //*其中：出口(千元)
            writeDataIntoCell(summarySheet, 21, 10, summaryCompanyRDData.NewProductSaleOfOutlet);

            //(三)其他情况
            //*期末拥有注册商标(件)
            writeDataIntoCell(summarySheet, 23, 10, summaryCompanyRDData.TrademarkOfPeriod);
            //发表科技论文(篇)
            writeDataIntoCell(summarySheet, 24, 10, summaryCompanyRDData.ScientificPapers);
            //形成国家或行业标准(项)
            writeDataIntoCell(summarySheet, 25, 10, summaryCompanyRDData.StandardsOfNational);

            //八、其他相关情况
            //(一)技术改造和技术获取情况
            //技术改造经费支出（千元）
            writeDataIntoCell(summarySheet, 28, 10, summaryCompanyRDData.TechTransformExpenses);
            //购买境内技术经费支出（千元）
            writeDataIntoCell(summarySheet, 29, 10, summaryCompanyRDData.BuyDomesticTechExpenses);
            //引进境外技术经费支出（千元）
            writeDataIntoCell(summarySheet, 30, 10, summaryCompanyRDData.ImpOverseasTechExpenses);
            //引进境外技术的消化吸收经费支出（千元）
            writeDataIntoCell(summarySheet, 31, 10, summaryCompanyRDData.ImpOverseasTechDigestionExpenses);

            // (二)企业办研究开发机构（境外）情况
            //期末企业在境外设立的研究开发机构数(个)
            writeDataIntoCell(summarySheet, 33, 10, summaryCompanyRDData.OverseasOrgCount);


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

        //获取单元格的值
        public string getCellValueByCellType(ISheet sheet, int rowIndex, int colIndex)
        {

            //获取到单元格
            ICell cell = sheet.GetRow(rowIndex).GetCell(colIndex);
            // 判断单元格是否存在  
            if (cell != null)
            {
                // 判断单元格类型  
                if (cell.CellType == CellType.String)
                {
                    // 字符串类型  
                    return cell.StringCellValue;
                }
                else if (cell.CellType == CellType.Numeric)
                {
                    // 数字类型  
                    return cell.NumericCellValue.ToString();
                }
                else if (cell.CellType == CellType.Formula)
                {
                    // 公式类型，需要计算后获取值  
                    return cell.StringCellValue;
                }
                else if (cell.CellType == CellType.Blank)
                {
                    // 空白类型，没有值  
                    return "";
                }
                else if (cell.CellType == CellType.Error)
                {
                    // 错误类型，需要处理错误情况  
                    return ""; // 需要根据实际情况处理错误值  
                }
                else
                {
                    //TODO 扔异常，父级拿到异常后，加入到错误日志中
                    return "";
                }
            }
            else {
                return "";
            }
        }

        //正则获取到数字返回
        public float getFloat(string numberStr) {

            //通过正则获取到表格中的数据
            string pattern = @"^\d+(\.\d+)?$"; // 匹配一串连续的数字  

            Regex regex = new Regex(pattern);
            System.Text.RegularExpressions.Match match = regex.Match(numberStr);

            if (match.Success)
            {
                return float.Parse(match.Value);
            }
            else
            {
                return 0;
            }
        }

        //正则获取到数字返回
        public int getInt(string numberStr)
        {

            //通过正则获取到表格中的数据
            string pattern = @"^\d+(\.\d+)?$"; // 匹配一串连续的数字  

            Regex regex = new Regex(pattern);
            System.Text.RegularExpressions.Match match = regex.Match(numberStr);

            if (match.Success)
            {
                return int.Parse(match.Value);
            }
            else
            {
                return 0;
            }
        }

        //获取单元格的值
        public void writeDataIntoCell(ISheet sheet, int rowIndex, int colIndex, dynamic cellValue)
        {
            sheet.GetRow(rowIndex).GetCell(colIndex).SetCellValue(cellValue);
        }

        //删除指定文件夹下的文件
        static void DeleteAllFiles(string folderPath)
        {
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            FileInfo[] files = directory.GetFiles();

            foreach (FileInfo file in files)
            {
                if (file.Exists)
                {
                    file.Delete();
                }
            }
        }
    }
}
