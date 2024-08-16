using ConsoleApp1.Models;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using SQLite;
using System.Collections.ObjectModel;

namespace ConsoleApp1.Persistence;
internal static class ApplicationDbContext
{
    public static SQLiteConnection db;
    static ApplicationDbContext()
    {
        //File.Delete(Path.Combine("D:\\work\\Univer\\Task 1 - Comments of modules (read word and paste into excel)\\ConsoleApp1\\", @"consoleDb.db"));

    }

    public static void Init()
    {
        if (db != null) return;
        var path = Path.Combine(ConsoleApp1.Helpers.SettingsHelper.GetProjectPath(), @"consoleDb.db");
        db = new SQLiteConnection(path);
        db.CreateTable<Module>();
    }

    public static void Insert(object obj)
    {
        Init();
        if (obj == null) return;
        db.Insert(obj);
    }

    public static void Remove(object obj)
    {
        Init();
        if (obj == null) return;

        db.Delete(obj);

    }
    public static void Update(object obj)
    {
        Init();
        if (obj == null) return;

        db.Update(obj);

    }
    public static T Select<T>(int id) where T : new()
    {
        Init();
        return db.Get<T>(id);

    }
    public static ObservableCollection<T> SelectAll<T>() where T : new()
    {
        Init();
        return new ObservableCollection<T>(db.Table<T>());

    }

    
}
