namespace ConsoleApp1.Executers.Word;

public static class ModuleWordHelper
{
    public static bool IsDescriptionRow(string rowCellDescr) => rowCellDescr.Contains("кратк") && rowCellDescr.Contains("содерж")
     || rowCellDescr.Contains("змест") && rowCellDescr.Contains("дысцыпл");

    public static bool IsNameRow(string rowCellDescr) => rowCellDescr.Contains("названи") && rowCellDescr.Contains("дисциплин")
        || rowCellDescr.Contains("назва") && rowCellDescr.Contains("дысцыпл");

    public static bool IsSpecialityRow(string rowCellDescr) => (rowCellDescr.Contains("код") || rowCellDescr.Contains("назван")) && rowCellDescr.Contains("специальн")
        || rowCellDescr.Contains("код") && rowCellDescr.Contains("спецыяльнасц");
}
