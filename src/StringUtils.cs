namespace SoftUniQuizGenerator
{
	public static class StringUtils
	{
		public static string TruncateString(string? input, int maxLength = 150)
		{
			if (input == null)
				input = string.Empty;
			input = input.Replace("\r", " ").Trim();
			if (input.Length <= maxLength)
			{
				return input;
			}
			else if (maxLength <= 3)
			{
				return input.Substring(0, maxLength);
			}
			else
			{
				int midLength = maxLength - 3;
				int halfLength = midLength / 2;
				string start = input.Substring(0, halfLength);
				string end = input.Substring(input.Length - halfLength);
				return $"{start}...{end}";
			}
		}
	}
}
