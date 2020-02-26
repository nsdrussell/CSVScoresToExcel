using System;
using System.Linq;

namespace ScoresToExcelApp
{
    internal static class ArithmaticHelpers
    {
        /// <summary>
        /// The tolerance of scores allowed
        /// </summary>
        private const double PercentTolerance = 0.2d;

        /// <summary>
        /// Computes the truncated (trimmed) mean of the given values.
        /// </summary>
        /// <param name="scores">The scores to get a truncated mean of.</param>
        public static double TruncatedMean(this string[] scores)
        {
            double[] doubleScores = scores.Select(score => double.Parse(score)).ToArray();
            Array.Sort(doubleScores);

            int k = (int)Math.Floor(doubleScores.Length * PercentTolerance);

            double sum = 0;
            for (int i = k; i < doubleScores.Length - k; i++)
                sum += doubleScores[i];

            return sum / (doubleScores.Length - 2 * k);
        }

        internal static int[] GetErroneousScoresGivenTruncatedMean(int[] scores, double truncatedMean)
        {
            return scores.Where(score => score < truncatedMean - 10 || score > truncatedMean + 10).ToArray();
        }
    }
}