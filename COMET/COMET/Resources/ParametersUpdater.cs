using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace COMET.Resources
{
    public class ParametersUpdater
    {
        public class Parameter
        {
            public string Name { get; set; }
            public string Type { get; set; }
            public List<string> Description { get; set; } = new List<string>();
        }

        /// <summary>
        /// Updates the XML documentation comment's Parameters tag based on the provided current and updated parameter texts.
        /// Assumes input strings are single-line with entries prefixed by '///', representing parameters and description lines.
        /// </summary>
        /// <param name="currentParams">Text for current parameters, e.g., "/// sender(object): the text/// more text/// e(EventArgs): thee text/// more text".</param>
        /// <param name="updatedParams">Text for updated parameters, e.g., "/// param1(int): /// param2(string):".</param>
        /// <param name="indent">The indentation string to use for formatting the output.</param>
        /// <returns>The complete updated <Parameters> tag with newlines added.</returns>
        public static string UpdateParameters(string currentParams, string updatedParams, string indent)
        {
            List<Parameter> ParseParameters(string paramsText)
            {
                var parameters = new List<Parameter>();
                if (string.IsNullOrWhiteSpace(paramsText))
                    return parameters;

                // Split by '///' and remove empty entries
                var entries = paramsText.Split(new[] { "///" }, StringSplitOptions.RemoveEmptyEntries);
                Parameter currentParam = null;

                foreach (var entry in entries)
                {
                    var trimmedEntry = entry.Trim();
                    if (string.IsNullOrEmpty(trimmedEntry))
                        continue;

                    // Check if the entry is a parameter (matches name(type): pattern)
                    var match = Regex.Match(trimmedEntry, @"^(\w+)\(([^()]+)\):(?:\s*(.*))?");
                    if (match.Success)
                    {
                        if (currentParam != null)
                            parameters.Add(currentParam);

                        currentParam = new Parameter
                        {
                            Name = match.Groups[1].Value,
                            Type = match.Groups[2].Value
                        };
                        if (!string.IsNullOrEmpty(match.Groups[3].Value))
                            currentParam.Description.Add(match.Groups[3].Value);
                    }
                    else if (currentParam != null)
                    {
                        // Non-parameter entry is a description line
                        currentParam.Description.Add(trimmedEntry);
                    }
                }

                if (currentParam != null)
                    parameters.Add(currentParam);

                return parameters;
            }

            string BuildParameters(List<Parameter> parameters, List<Parameter> referenceParams, string indents)
            {
                var sb = new StringBuilder();
                for (int i = 0; i < parameters.Count; i++)
                {
                    var param = parameters[i];
                    var desc = (referenceParams != null && i < referenceParams.Count) ? referenceParams[i].Description : param.Description;
                    sb.Append($"{indents}///    {param.Name}({param.Type}):");
                    if (desc.Count > 0)
                        sb.Append($" {desc[0]}");
                    sb.AppendLine(); // Add newline after parameter line

                    // Add additional description lines
                    for (int j = 1; j < desc.Count; j++)
                        sb.AppendLine($"{indents}///    {desc[j]}");
                }
                return sb.ToString().TrimEnd('\n');
            }

            var currentParamList = ParseParameters(currentParams);
            var updatedParamList = ParseParameters(updatedParams);

            if (updatedParamList.Count <= currentParamList.Count)
            {
                // Update names and types, transfer descriptions only for matching parameter names
                foreach (var updatedParam in updatedParamList)
                {
                    // Find a current parameter with the same name
                    var matchingCurrentParam = currentParamList.FirstOrDefault(p => p.Name == updatedParam.Name);
                    if (matchingCurrentParam != null)
                    {
                        // Transfer the description from the matching current parameter
                        updatedParam.Description = matchingCurrentParam.Description;
                    }
                    // If no match, updatedParam.Description remains as is (empty or from updatedParams)
                }
            }
            else
            {
                // Copy descriptions for matching parameters, leave new ones as is
                foreach (var updatedParam in updatedParamList)
                {
                    // Find a current parameter with the same name
                    var matchingCurrentParam = currentParamList.FirstOrDefault(p => p.Name == updatedParam.Name);
                    if (matchingCurrentParam != null)
                    {
                        // Transfer the description from the matching current parameter
                        updatedParam.Description = matchingCurrentParam.Description;
                    }
                    // New parameters or non-matching names keep their descriptions (empty or from updatedParams)
                }
            }

            return $"/// <Parameters>\n{BuildParameters(updatedParamList, null, indent)}\n{indent}/// </Parameters>\n";
        }
    }
}