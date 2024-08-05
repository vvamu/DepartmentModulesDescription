namespace ConsoleApp1.Models;

public class Module
{
    public string DepartmentShortName { get; set; } = " - ";
    public string Name { get; set; } = "Не задано или названо некорректно"; //if (node consists 'Название учебной дисциплины')
    public string Description { get; set; } = "Не задано"; //if (node consists 'Краткое содержание учебной дисциплины')

    public DateTime DateLastUpdateFile {get;set;}



    public override bool Equals(object obj)
    {
        if (obj == null || GetType() != obj.GetType())
            return false;

        Module other = (Module)obj;
        return Name == other.Name && DepartmentShortName == other.DepartmentShortName && Description == other.Description;
    }


}
