namespace Tabel;

internal enum FuzzyMatchStatus
{
    NotFound,
    Found,
    Ambiguous
}

internal sealed record FuzzyMatchResult<T>(
    FuzzyMatchStatus Status,
    T? Match,
    IReadOnlyList<T> Candidates,
    int Distance,
    double DifferencePercent);

internal static class FioFuzzyMatcher
{
    private const double DifferenceThreshold = 0.05d;

    public static FuzzyMatchResult<T> Find<T>(
        string normalizedSource,
        IEnumerable<T> candidates,
        Func<T, string> normalizedNameSelector)
    {
        var eligible = new List<(T Candidate, int Distance, double Percent)>();

        foreach (T candidate in candidates)
        {
            string candidateName = normalizedNameSelector(candidate);
            if (string.IsNullOrWhiteSpace(candidateName))
                continue;

            int significantLength = Math.Max(
                CountSignificantCharacters(normalizedSource),
                CountSignificantCharacters(candidateName));
            int allowedDistance = Math.Max(1, (int)Math.Floor(significantLength * DifferenceThreshold));
            int distance = LevenshteinDistance(normalizedSource, candidateName, allowedDistance);

            if (distance <= allowedDistance)
            {
                double percent = significantLength == 0
                    ? 0d
                    : (double)distance / significantLength * 100d;
                eligible.Add((candidate, distance, percent));
            }
        }

        if (eligible.Count == 0)
            return new FuzzyMatchResult<T>(FuzzyMatchStatus.NotFound, default, [], 0, 0d);

        int bestDistance = eligible.Min(x => x.Distance);
        var best = eligible.Where(x => x.Distance == bestDistance).ToList();

        if (best.Count == 1)
        {
            return new FuzzyMatchResult<T>(
                FuzzyMatchStatus.Found,
                best[0].Candidate,
                [best[0].Candidate],
                bestDistance,
                best[0].Percent);
        }

        return new FuzzyMatchResult<T>(
            FuzzyMatchStatus.Ambiguous,
            default,
            best.Select(x => x.Candidate).ToList(),
            bestDistance,
            best.Min(x => x.Percent));
    }

    private static int CountSignificantCharacters(string value)
    {
        return value.Count(c => !char.IsWhiteSpace(c));
    }

    private static int LevenshteinDistance(string source, string target, int maximumDistance)
    {
        if (Math.Abs(source.Length - target.Length) > maximumDistance)
            return maximumDistance + 1;
        if (source.Length == 0)
            return target.Length;
        if (target.Length == 0)
            return source.Length;

        var previous = new int[target.Length + 1];
        var current = new int[target.Length + 1];
        int outsideLimit = maximumDistance + 1;

        for (int column = 0; column <= target.Length; column++)
            previous[column] = column <= maximumDistance ? column : outsideLimit;

        for (int row = 1; row <= source.Length; row++)
        {
            Array.Fill(current, outsideLimit);
            if (row <= maximumDistance)
                current[0] = row;

            int firstColumn = Math.Max(1, row - maximumDistance);
            int lastColumn = Math.Min(target.Length, row + maximumDistance);

            for (int column = firstColumn; column <= lastColumn; column++)
            {
                int substitutionCost = source[row - 1] == target[column - 1] ? 0 : 1;
                current[column] = Math.Min(
                    Math.Min(current[column - 1] + 1, previous[column] + 1),
                    previous[column - 1] + substitutionCost);
            }

            (previous, current) = (current, previous);
        }

        return previous[target.Length] <= maximumDistance
            ? previous[target.Length]
            : outsideLimit;
    }
}
