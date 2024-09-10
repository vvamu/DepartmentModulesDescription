using ConsoleApp1.Models;
using SQLite;
using System.ComponentModel.DataAnnotations.Schema;

namespace ConsoleApp1.Models;

public class ModuleWrite
{
    public string Name { get; set; } = string.Empty;
    public string Exams { get; set; }

    public string Receives { get; set; } = string.Empty; public string GetReceives => Receives.ToLower().Contains("д") ? Receives.ToLower().Replace("д", "(дифференцированный зачет)") : Receives;
    public string TotalHours { get; set; } = string.Empty;

    public string AuditoriumHours { get; set; } = string.Empty;
    public string GetAuditoriumHours =>
       string.IsNullOrEmpty(AuditoriumHours) ? "" :
        (AuditoriumHours + " ауд. ч." + ((string.IsNullOrEmpty(LectureHours) && string.IsNullOrEmpty(LabsHours) && string.IsNullOrEmpty(PracticeHours) && string.IsNullOrEmpty(SeminarHours)) ? "" : ","));
    public string LectureHours { get; set; } = string.Empty;
    public string GetLectureHours =>
        LectureHours == null ? "" :
        (LectureHours + " лекционных ч." + ((string.IsNullOrEmpty(LabsHours) && string.IsNullOrEmpty(PracticeHours) && string.IsNullOrEmpty(SeminarHours)) ? "" : ","));

    public string LabsHours { get; set; } = string.Empty;
    public string GetLabsHours =>
        LabsHours == null ? "" :
        (LabsHours + " лаб. ч." + ((string.IsNullOrEmpty(PracticeHours) && string.IsNullOrEmpty(SeminarHours)) ? "" : ","));

    public string PracticeHours { get; set; } = string.Empty; 
    public string GetPracticeHours => PracticeHours == null ? "" : (PracticeHours + " практических ч." + (string.IsNullOrEmpty(SeminarHours) ? "" : ","));
    public string SeminarHours { get; set; } = string.Empty; 
    public string GetSeminarHours => SeminarHours == null ? "" : SeminarHours + " семинарских ч.";

    public string GetResultHours
    {
        get
        {
            var result = $"({GetAuditoriumHours} {GetLectureHours} {GetLabsHours} {GetPracticeHours} {GetSeminarHours})";
            if(string.IsNullOrEmpty(result.Replace(" ", "").Trim())) return "";
            result = result.Replace("  ", " ").Replace(" ,", ",").Replace(", ", ", ").Replace(" )", ")").Replace("( ", "(");
            return result;
        }
    }

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
