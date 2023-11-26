using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools
{
    internal class ProjectRDData
    {
        //项目名称
        public String RDProjectName { get; set; }
       
        //项目来源
        public String RDProjectSource { get; set; }
       
        //项目开展形式
        public String RDProjectDevForm { get; set; }
        
        //项目当年成果形式(Project current results form)
        public String RDProjectCurrentResultsForm { get; set; }

        //项目技术经济目标
        public String RDProjectEconomicTarget { get; set; }


        //项目起始日期
        public DateTime RDProjectBeginDate { get; set; }

        //项目完成日期
        public DateTime RDProjectEndDate { get; set; }

        //跨年项目当年所处主要进展阶段
        public String AcrossYearRDProjectCurrentStage { get; set; }

        //项目研究开发人员 （人）
        public int RDProjectResearcherCount { get; set; }

        //项目人员实际工作时间  （人月）
        public int RDProjectStaffWorkMonth { get; set; }

        //项目经费支出（千元）
        public decimal RDProjectExpenses { get; set; }

        //其中：政府资金
        public decimal RDProjectExpensesFromGovernment { get; set; }

        //*其中：用于科学原理的探索发现
        public decimal RDProjectExpensesForSicResearch { get; set; }

        //*其中：企业自主开展
        public decimal RDProjectExpensesFromComSelf { get; set; }

        //*委托外单位开展
        public decimal RDProjectExpensesFromEntrustOutsource { get; set; }

    }
}
