using SQLite;
using System.ComponentModel.DataAnnotations.Schema;

namespace ConsoleApp1.Models;

public class Module
{
    [PrimaryKey]
    public Guid Id { get; set; } = Guid.NewGuid();
    public string DepartmentShortName { get; set; }
    public string Speciality { get; set; }
    
    public string LowerAndTrimSpeciality { get; set; }

    public string Name { get; set; }

    public string Description { get; set; } 

    public string DateLastUpdateFileString {get;set;}
    [NotMapped]
    public bool IsDocxConvertedByDoc { get; set; }

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
            ;
    }

    public bool IsDifferentDescriptionAndDateLastUpdate(Module other)
    {
        return Description != other.Description && DateLastUpdateFileString != other.DateLastUpdateFileString;
    }
    public Module() { }


    public Module(string departmentShortName,string speciality, string name, string description,string fullFilePath = "", string fileName = "", bool isDocxConvertedByDoc = false)
    {
        DepartmentShortName = departmentShortName;
        Speciality = speciality;
        Name = name;
        Description = description;
        IsDocxConvertedByDoc = isDocxConvertedByDoc;
        if (!string.IsNullOrEmpty(fullFilePath)) FullFilePath = fullFilePath;
        if (!string.IsNullOrEmpty(fileName)) FileName = fileName;
    }
    public Module(string departmentShortName, string fileName, string fullFilePath, bool isDocxConvertedByDoc)
    {
        DepartmentShortName = departmentShortName;
        FileName = fileName;
        FullFilePath = fullFilePath;
        IsDocxConvertedByDoc = isDocxConvertedByDoc;
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
