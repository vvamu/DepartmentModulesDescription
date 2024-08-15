using System.ComponentModel.DataAnnotations.Schema;

namespace ConsoleApp1.Models;

public class Module
{
    public string DepartmentShortName { get; set; }
    public string Speciality { get; set; }

    public string Name { get; set; } //= "Не задано или названо некорректно"; //if (node consists 'Название учебной дисциплины')

    public string Description { get; set; } //= "Не задано"; //if (node consists 'Краткое содержание учебной дисциплины')

    public string DateLastUpdateFileString {get;set;}
    //[NotMapped]
    //public DateTime DateLastUpdateFile { get; set; }

    public string FileName { get; set; }
    public string FullFilePath { get; set; } 




    public override string ToString()
    {
        return $"{DepartmentShortName} - {FileName}; Speciality {Speciality}; Module: {Name} - {Description}  ";
    }
    public override bool Equals(object obj)
    {
        if (obj == null || GetType() != obj.GetType())
            return false;

        Module other = (Module)obj;
        return Speciality == other.Speciality 
            &&  Name == other.Name 
            && DepartmentShortName == other.DepartmentShortName 
            && Description == other.Description
            && (DateLastUpdateFileString == other.DateLastUpdateFileString);
    }

    /*
    public bool Compare(object obj)
    {
        if (obj == null || GetType() != obj.GetType())
            return false;

        Module other = (Module)obj;
        return GetHashCode() == other.GetHashCode();
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Speciality, Name, DepartmentShortName, Description, DateLastUpdateFile);
    }
    */

}
