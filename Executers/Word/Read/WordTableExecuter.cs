using DocumentFormat.OpenXml.Wordprocessing;
using ConsoleApp1.Models;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using ConsoleApp1.Executers.Word.Helpers;

namespace ConsoleApp1.Helpers;
public partial class WordFileReader
{
    private async Task<(Module, List<Module>)> ProcessTable(Table node, Module resultModule)
    {

        List<Module> modules = new List<Module>();
        var isFoundOneTable = 0;
        var module = new Module();
        foreach (var row in node.Descendants<TableRow>())
        {
            Module? rowModule;
            ///TODO
            try
            {
                rowModule = ProcessRowTopDown(row);
            }
            catch (Exception ex)
            {
                rowModule = ProcessRowLeftRight(row);
            }

            if (rowModule == null) continue;
            if (ModuleWordHelper.IsSpecialityRow(row.InnerText.ToLower().Trim())) isFoundOneTable++;
            if (isFoundOneTable > 1) { modules.Add(module); module = new Module(resultModule.DepartmentShortName, resultModule.FileName, resultModule.FullFilePath, resultModule.IsDocxConvertedByDoc); isFoundOneTable = 1; }

            if (!string.IsNullOrEmpty(rowModule.Speciality)) module.Speciality = rowModule.Speciality;
            if (!string.IsNullOrEmpty(rowModule.Name)) module.Name = rowModule.Name;
            if (!string.IsNullOrEmpty(rowModule.Description)) module.Description = rowModule.Description;
        }
        resultModule.Speciality = modules.FirstOrDefault()?.Speciality ?? module.Speciality;
        resultModule.Name = modules.FirstOrDefault()?.Name ?? module.Name;
        resultModule.Description = modules.FirstOrDefault()?.Description ?? module.Description;


        if (this.IsNeedToTranslate)
        {
            await TranslationHelper.TranslateToRu(resultModule);
        }
        return (resultModule, modules);
    }
    private Module ProcessRowTopDown(TableRow row)
    {
        var module = new Module();
        var rowText = row.InnerText.ToLower().Replace(" ", "").Replace("-", "");
        if (ModuleWordHelper.IsDescriptionRow(rowText) || ModuleWordHelper.IsNameRow(rowText) || ModuleWordHelper.IsSpecialityRow(rowText))
        {
            var rowArray = row.Descendants<TableCell>().ToArray();

            for (int i = 0; i < row.Descendants<TableCell>().Count() - 1; i++)
            {
                var rowCellDescr = rowArray[i].InnerText.ToLower().Replace(" ", "").Replace("-", "");
                if (ModuleWordHelper.IsDescriptionRow(rowCellDescr))
                {
                    module.Description = rowArray[i + 1].InnerText;

                }
                else if (ModuleWordHelper.IsNameRow(rowCellDescr))
                {
                    module.Name = rowArray[i + 1].InnerText;
                }
                else if (ModuleWordHelper.IsSpecialityRow(rowCellDescr))
                {
                    module.Speciality = rowArray[i + 1].InnerText;
                }
            }
            return module;
        }
        return null;

    }
    //TODO
    private Module ProcessRowLeftRight(TableRow row)
    {
        return new Module();
    }

}
