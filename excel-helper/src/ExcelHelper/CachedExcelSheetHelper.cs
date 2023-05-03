using HP.SV.DotNetRuleApi;
using System.IO;

namespace OpenText.ExcelHelper {
    public static class CachedExcelSheetHelper {

        public static ExcelSheet GetOrCreateCachedExcelSheet(this HpsvPersistentContext context, string filePath, string sheetName, bool hasHeader, string lookupColumn) {

            // check cache validity
            long lastModificationTimeTicks = File.GetLastWriteTime(filePath).Ticks;

            string CachedExcelChangeTimestampKey = $"ExcelChangedTimestampKey-[{filePath}]";
            string CachedExcelSheetKey = $"SheetKey-[{filePath}][{sheetName}][{hasHeader}][{lookupColumn}]";

            long cachedExcelChangeTimestampTicks;
            context.TryGetValue(CachedExcelChangeTimestampKey, out cachedExcelChangeTimestampTicks);

            if (cachedExcelChangeTimestampTicks < lastModificationTimeTicks || !context.ContainsKey(CachedExcelSheetKey)) {

                ExcelSheet sheet = new ExcelSheet(filePath, sheetName, hasHeader, lookupColumn);

                context.Add(CachedExcelChangeTimestampKey, lastModificationTimeTicks);
                context.Add(CachedExcelSheetKey, sheet);
            }
            return (ExcelSheet)context[CachedExcelSheetKey];

        }

    }
}
