using HP.SV.DotNetRuleApi;
using System.Collections.Generic;

namespace OpenText.EvaluateVariablesExtension {

    public static class EvaluateVariablesExtension {

        /// <summary>
        /// Replaces %%variableNames%% with the actual values, which must be supplied in variables dictionary.
        /// </summary>
        /// <param name="svObject">Node from which to start replacement</param>
        /// <param name="variables">Dictionary which maps variable names to its values</param>
        /// <param name="logger">Optional logger. Used for e.g. telling if some variable name is missing in the dictionary</param>
        public static void EvaluateVariables(this HpsvObject svObject, Dictionary<string, string> variables, HpsvLogger logger = null) {
            Traverse(svObject, variables, "%%", "%%", logger);
        }

        /// <summary>
        /// Replaces {variableNameStart}variableNames{variableNameEnd} with the actual values, which must be supplied in variables dictionary.
        /// </summary>
        /// <param name="svObject">Node from which to start replacement</param>
        /// <param name="variables">Dictionary which maps variable names to its values</param>
        /// <param name="logger">Optional logger. Used for e.g. telling if some variable name is missing in the dictionary</param>
        /// <param name="variableNamePrefix">Characters prefixing the actual variable name. e.g. "{"</param>
        /// <param name="variableNameSuffix">Characters suffixing the actual variable name. e.g. "}"</param>
        public static void EvaluateVariables(this HpsvObject svObject, Dictionary<string, string> variables, string variableNamePrefix, string variableNameSuffix, HpsvLogger logger = null) {
            Traverse(svObject, variables, variableNamePrefix, variableNameSuffix, logger);
        }

        private static void Traverse(HpsvObject svObject, Dictionary<string, string> variables, string variableNameStart, string variableNameEnd, HpsvLogger logger) {
            if (svObject == null) {
                return;
            }

            IEnumerable<KeyValuePair<string, object>> properties = svObject.GetAllProperties();
            foreach (KeyValuePair<string, object> property in properties) {
                string key = property.Key;
                object value = property.Value;

                if (value == null) {
                    continue;
                }

                if (value.GetType().IsSubclassOf(typeof(HpsvObject))) {
                    Traverse((HpsvObject)value, variables, variableNameStart, variableNameEnd, logger);
                } else if (value.GetType().IsSubclassOf(typeof(HpsvArray))) {
                    Traverse((HpsvArray)value, variables, variableNameStart, variableNameEnd, logger);

                } else if (typeof(string) == value.GetType()) {
                    string stringValue = (string)value;
                    string valueOut;
                    if (ProcessValue(stringValue, out valueOut, variables, variableNameStart, variableNameEnd, logger)) {
                        svObject[key] = valueOut;
                    }
                }
            }
        }

        private static void Traverse(HpsvArray svArray, Dictionary<string, string> variables, string variableNameStart, string variableNameEnd, HpsvLogger logger) {
            for (int i = 0; i < svArray.Count; i++) {
                object arrayItem = svArray.GetGenericItem(i);

                if (arrayItem.GetType().IsSubclassOf(typeof(HpsvObject))) {
                    Traverse((HpsvObject)arrayItem, variables, variableNameStart, variableNameEnd, logger);
                } else if (arrayItem.GetType().IsSubclassOf(typeof(HpsvArray))) {
                    Traverse((HpsvArray)arrayItem, variables, variableNameStart, variableNameEnd, logger);

                } else if (typeof(string) == arrayItem.GetType()) {
                    string stringValue = (string)arrayItem;
                    string valueOut;
                    if (ProcessValue(stringValue, out valueOut, variables, variableNameStart, variableNameEnd, logger)) {
                        svArray.SetGenericItem(i, valueOut);
                    }
                }
            }
        }

        private static bool ProcessValue(string valueIn, out string valueOut, Dictionary<string, string> variables, string variableNameStart, string variableNameEnd, HpsvLogger logger) {
            valueOut = null;

            if (valueIn.StartsWith(variableNameStart) && valueIn.EndsWith(variableNameEnd)) {

                // is this step necessary? maybe I got just wrong message samples, which suffered from HTML entities conversion...
                valueIn = valueIn.Substring(2, valueIn.Length - 4);

                if (variables.TryGetValue(valueIn, out valueOut)) {
                    return true;
                } else {
                    if (logger != null) {
                        logger.Error($"Unable to find variable \"{valueIn}\"");
                    }
                }
            }
            return false;
        }


    }
}
