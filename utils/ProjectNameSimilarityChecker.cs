using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools.utils
{
    public static class ProjectNameSimilarityChecker
    {
        /// <summary>
        /// 计算莱文斯坦相似度（编辑距离）
        /// 相似度 = 1 - 编辑距离 / 最大长度
        /// </summary>
        public static double CalculateLevenshteinSimilarity(string a, string b)
        {
            if (string.IsNullOrEmpty(a))
                return string.IsNullOrEmpty(b) ? 1.0 : 0.0;
            if (string.IsNullOrEmpty(b))
                return 0.0;

            int maxLength = Math.Max(a.Length, b.Length);
            int distance = ComputeLevenshteinDistance(a, b);
            return 1.0 - (double)distance / maxLength;
        }

        private static int ComputeLevenshteinDistance(string a, string b)
        {
            int[,] matrix = new int[a.Length + 1, b.Length + 1];

            for (int i = 0; i <= a.Length; i++) matrix[i, 0] = i;
            for (int j = 0; j <= b.Length; j++) matrix[0, j] = j;

            for (int i = 1; i <= a.Length; i++)
            {
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            }
            return matrix[a.Length, b.Length];
        }

        /// <summary>
        /// 计算N元语法Jaccard相似度
        /// 相似度 = 交集大小 / 并集大小
        /// </summary>
        public static double CalculateJaccardSimilarity(string a, string b, int nGramSize = 2)
        {
            if (nGramSize <= 0)
                throw new ArgumentException("nGram大小必须为正整数");

            var nGramsA = GenerateNGrams(a, nGramSize);
            var nGramsB = GenerateNGrams(b, nGramSize);

            var setA = new HashSet<string>(nGramsA);
            var setB = new HashSet<string>(nGramsB);

            int intersection = setA.Intersect(setB).Count();
            int union = setA.Union(setB).Count();

            return union == 0 ? 0.0 : (double)intersection / union;
        }

        private static IEnumerable<string> GenerateNGrams(string str, int n)
        {
            if (string.IsNullOrEmpty(str) || str.Length < n)
                return Enumerable.Empty<string>();

            List<string> nGrams = new List<string>();
            for (int i = 0; i <= str.Length - n; i++)
                nGrams.Add(str.Substring(i, n));

            return nGrams;
        }

        /// <summary>
        /// 综合相似度计算（加权平均）
        /// </summary>
        public static double CalculateCombinedSimilarity(string a, string b, double levenWeight = 0.5)
        {
            double leven = CalculateLevenshteinSimilarity(a, b);
            double jaccard = CalculateJaccardSimilarity(a, b, 2);
            return leven * levenWeight + jaccard * (1 - levenWeight);
        }
    }
}
