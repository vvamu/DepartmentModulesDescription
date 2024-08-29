using SQLite;
using System.ComponentModel.DataAnnotations.Schema;

namespace ConsoleApp1.Models;

public class ModuleWrite
{
    public string Name { get; set; }
    public string Exams { get; set; }
    public string Receives { get; set; }
    public string TotalHours { get; set; }
    public string AuditoriumHours { get; set; }
    public string LectureHours { get; set; }
    public string LabsHours { get; set; }

    public string PracticeHours { get; set; }
    public string SeminarHours { get; set; }
    public string DepartmentShortName { get; set; }

    public string ReceivedUnits { get; set; }

    public string Description { get; set; }
    public string Speciality { get; set; }

    public string FileName { get; set; }

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

}
