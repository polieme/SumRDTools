using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools
{
    internal class CompanyRDData
    {
        //一、研究开发人员情况
        // 研究开发人员合计(人)
        public int RDPersonnelTotal { get; set; }
        //其中：管理和服务人员(人)
        public int RDPersonnelManageAndService { get; set; }
        //其中：女性(人)
        public int RDPersonnelFemale { get; set; }
        //其中：全职人员(人)
        public int RDPersonnelFullTimeStaff { get; set; }
        //其中：本科毕业及以上人员(人)
        public int RDPersonnelBachelorAndAbove { get; set; }
        //其中：外聘人员(人)
        public int RDPersonnelExternalStaff { get; set; }

        //二、研究开发费用情况
        // 研究开发费用合计（千元）
        public decimal RDExpensesTotal { get; set; }
        //1.人员人工费用（千元）
        public decimal RDExpensesPersonnelLabor { get; set; }
        //2.直接投入费用（千元）
        public decimal RDExpensesDirectInput { get; set; }
        //3.折旧费用与长期待摊费用（千元）(Depreciation and long-term amortization expenses)
        public decimal RDExpensesDepreciationAndLongTerm { get; set; }
        //4.无形资产摊销费用（千元）(Amortization expense of intangible assets)
        public decimal RDExpensesIntangibleAssets { get; set; }
        //5.设计费用（千元）
        public decimal RDExpensesDesign { get; set; }
        //6.装备调试费用与试验费用（千元）(Equipment commissioning costs and test costs)
        public decimal RDExpensesEquipmentDebug { get; set; }
        //7.委托外部研究开发费用（千元）(Outsourced R&D)
        public decimal RDExpensesEntrustOutsourcedRD { get; set; }
        //①委托境内研究机构（千元）（Entrust domestic research institutions）
        public decimal RDExpensesEntrustDomesticResearch { get; set; }
        //②委托境内高等学校（千元）
        public decimal RDExpensesEntrustDomesticCollege { get; set; }
        //③委托境内企业（千元）
        public decimal RDExpensesEntrustDomesticCompany { get; set; }
        //④委托境外机构（千元）(Overseas Institutions)
        public decimal RDExpensesEntrustOverseasInstitutions { get; set; }
        //8.其他费用（千元）
        public decimal RDExpensesOthers { get; set; }


        //三、研究开发资产情况
        //当年形成用于研究开发的固定资产（千元）
        public decimal RDAssetsYear { get; set; }
        //其中：仪器和设备（千元）
        public decimal RDAssetsYearEquipment { get; set; }

        //四、研究开发支出资金来源(Sources of funding for R&D expenditures)
        //1.来自企业自筹(千元)
        public decimal RDSpendSourceOfCompany { get; set; }
        //2.来自政府部门（千元）
        public decimal RDSpendSourceOfGovernment { get; set; }
        //3.来自银行贷款（千元）
        public decimal RDSpendSourceOfBank { get; set; }
        //4.来自风险投资（千元）
        public decimal RDSpendSourceOfRiskCapital { get; set; }
        //5.来自其他渠道（千元）
        public decimal RDSpendSourceOfOthers { get; set; }

        //五、相关政策落实情况(Policy implementation)
        //申报加计扣除减免税的研究开发支出(千元)(Declare additional deduction and subtraction)
        public decimal PolicyImplementDeclareAddtionRD { get; set; }
        //加计扣除减免税金额(千元)
        public decimal PolicyImplementAddtionRDTaxFree { get; set; }
        //高新技术企业减免税金额(千元)（The amount of tax reduction and exemption for high-tech enterprises）
        public decimal PolicyImplementHighTechRDTaxFree { get; set; }

        //六、企业办研究开发机构（境内）情况(Enterprise-run research and development institutions (domestic).)
        //期末机构数(个)
        public int CompanyRunOrgCountEndOfPeriod { get; set; }
        //机构研究开发人员（人）
        public int CompanyRunOrgRDPersonnel { get; set; }
        //其中：博士毕业（人）
        public int CompanyRunOrgRDDoctor { get; set; }
        //其中：硕士毕业（人）
        public int CompanyRunOrgRDMaster { get; set; }
        //机构研究开发费用（千元）
        public decimal CompanyRunOrgRDExpenses { get; set; }
        //期末仪器和设备原价（千元）
        public decimal CompanyRunOrgEquipmentValueEndOfPeriod { get; set; }


        //七、研究开发产出及相关情况
        //(一) 专利情况
        //当年专利申请数(件)(The number of patent applications filed in the current year)
        public int PatentApplyOfCurrentYear { get; set; }
        //其中：发明专利（件）
        public int PatentApplyOfInvention { get; set; }
        //期末有效发明专利数（件）(The number of invention patents in force at the end of the period)
        public int PatentApplyOfInForcePeriod { get; set; }
        //其中：已被实施（件）
        public int PatentApplyOfBeenImplement { get; set; }
        //专利所有权转让及许可数（件）
        public int PatentApplyOfAssignment { get; set; }
        //专利所有权转让及许可收入（千元）
        public decimal PatentApplyOfAssignmentIncome { get; set; }

        //(二) 新产品情况
        //*新产品销售收入(千元)(New product sales revenue)
        public decimal NewProductSaleRevenue { get; set; }
        //*其中：出口(千元)
        public decimal NewProductSaleOfOutlet { get; set; }


        //(三)其他情况
        //*期末拥有注册商标(件)
        public int TrademarkOfPeriod { get; set; }
        //发表科技论文(篇)(Publication of scientific papers)
        public int ScientificPapers { get; set; }
        //形成国家或行业标准(项)(Form national or industry standards)
        public int StandardsOfNational { get; set; }


        //八、其他相关情况
        //(一)技术改造和技术获取情况
        //技术改造经费支出（千元）（Expenditure on technological transformation）
        public decimal TechTransformExpenses { get; set; }
        //购买境内技术经费支出（千元）（Expenditure on the purchase of domestic technology）
        public decimal BuyDomesticTechExpenses { get; set; }
        //引进境外技术经费支出（千元）
        public decimal ImpOverseasTechExpenses { get; set; }
        //引进境外技术的消化吸收经费支出（千元）
        public decimal ImpOverseasTechDigestionExpenses { get; set; }

        // (二)企业办研究开发机构（境外）情况
        //期末企业在境外设立的研究开发机构数(个)
        public int OverseasOrgCount { get; set; }

        //企业研发项目填报情况
        public List<ProjectRDData> projectRDDatas = new List<ProjectRDData>();

    }
}
