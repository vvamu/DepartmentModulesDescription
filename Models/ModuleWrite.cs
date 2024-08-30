using SQLite;
using System.ComponentModel.DataAnnotations.Schema;

namespace ConsoleApp1.Models;

public class ModuleWrite
{
    public string Name { get; set; } = string.Empty;
    public string Exams { get ; set; } 

    public string Receives { get; set; } = string.Empty;
    public string TotalHours { get; set; } = string.Empty;
    public string AuditoriumHours { get; set; } = string.Empty; public string GetAuditoriumHours => AuditoriumHours == null ? "0" : AuditoriumHours + " ауд. ч.";
    public string LectureHours { get; set; } = string.Empty; public string GetLectureHours => LectureHours == null ? "0" : LectureHours + " лекционных ч.";
    public string LabsHours { get; set; } = string.Empty; public string GetLabsHours => LabsHours == null ? "0" : LabsHours + "лаб. ч.";

    public string PracticeHours { get; set; } = string.Empty; public string GetPracticeHours => PracticeHours == null ? "0" : PracticeHours + " практических ч.";
    public string SeminarHours { get; set; } = string.Empty; public string GetSeminarHours => SeminarHours == null ? "0" : SeminarHours + " семинарских ч.";
    public string DepartmentShortName { get; set; } = string.Empty;

    public string ReceivedUnits { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;
    public string Speciality { get; set; } = string.Empty;

    public string FileName { get; set; } = string.Empty;
    public string FullPath { get; set; } = string.Empty;

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
