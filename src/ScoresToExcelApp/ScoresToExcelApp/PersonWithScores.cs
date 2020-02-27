using System;
using System.Linq;

namespace ScoresToExcelApp
{
    public class PersonWithScores
    {
        private const int MemberNameColumnIndex = 0;
        private const int MemberNumberColumnIndex = 1;
        private const int MemberUnadjustedAverageColumnIndex = 2;
        private const int MemberScoresColumnIndex = 3;

        public string MemberName { get; }
        public int MemberNumber { get; }

        public double UnadjustedAverage { get; }
        public double AdjustedAverage { get; private set; }

        public int[] Scores { get; }
        public int[] ErroneousScores { get; private set; }
        public int[] AdjustedScores { get; private set; }
        public double? PreviousAverage { get; private set; }

        public ScorerCategory Category { get; }
        public string ReadableCategory { get; }

        public string ScoresList
        {
            get { return string.Join(", ", Scores); }
        }

        public string ErroneousScoresList
        {
            get { return string.Join(", ", ErroneousScores); }
        }

        public PersonWithScores(string rowFromCSV)
        {
            //a line from the file looks like the following
            //"FirstName SecondName",1234,87.57,87,81,90,88,89,89,89

            var splitRow = rowFromCSV.Split(',');
            MemberName = splitRow[MemberNameColumnIndex].Trim('"');
            MemberNumber = int.Parse(splitRow[MemberNumberColumnIndex]);
            UnadjustedAverage = double.Parse(splitRow[MemberUnadjustedAverageColumnIndex]);
            //everything after third column are scores.

            var unparsedScores = new string[splitRow.Length - MemberScoresColumnIndex];
            Scores = new int[splitRow.Length - MemberScoresColumnIndex];

            for (int i = 0; i < Scores.Length; i++)
            {
                unparsedScores[i] = splitRow[i + MemberScoresColumnIndex];
                Scores[i] = int.Parse(unparsedScores[i]);
            }

            SetAdjustedAverageAndErroneousScores(unparsedScores);

            if (AdjustedScores.Length < 4)
            {
                Category = ScorerCategory.FewerThan4Cards;
                ReadableCategory = "Fewer Than 4 Cards";
            }
            else if (AdjustedAverage >= 90d)
            {
                Category = ScorerCategory.ClassA;
                ReadableCategory = "Class A";
            }
            else if (AdjustedAverage >= 80d)
            {
                Category = ScorerCategory.ClassB;
                ReadableCategory = "Class B";
            }
            else
            {
                Category = ScorerCategory.ClassC;
                ReadableCategory = "Class C";
            }
        }

        public void SetAdjustedAverageAndErroneousScores(string[] unParsedScores)
        {
            double truncatedMean = unParsedScores.TruncatedMean();

            ErroneousScores = ArithmaticHelpers.GetErroneousScoresGivenTruncatedMean(Scores, truncatedMean);

            AdjustedScores = Scores.Where(score => !ErroneousScores.Contains(score)).ToArray();

            AdjustedAverage = Math.Round(AdjustedScores.Average(), 2);
        }

        public void SetPreviousAverage(double previousAverage)
        {
            PreviousAverage = previousAverage;
        }

        public override string ToString()
        {
            return MemberName;
        }
    }
}