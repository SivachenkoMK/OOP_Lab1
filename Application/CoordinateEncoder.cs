using System;

namespace Application
{

    public static class CoordinateEncoder
    {
        public static string Encode(int x)
        {
            x++;
            var columnName = "";
            if (x == 0) return ((char)64).ToString();
            while (x > 0)
            {
                var mod = (x - 1) % 26;
                columnName = (char)(65 + mod) + columnName;
                x = (x - mod) / 26;
            }
            return columnName;
        }

        public static Tuple<int, int> Decode(string input)
        {
            var col = 0;
            var row = 0;
            foreach (var symbol in input)
            {
                switch (symbol)
                {
                    case >= 'A' and <= 'Z':
                        col *= 26;
                        col += (symbol - 64);
                        break;
                    case >= '0' and <= '9':
                        row *= 10;
                        row += (symbol - 48);
                        break;
                }
            }
            col -= 1;
            return new Tuple<int, int>(row, col);
        }
    }
}
