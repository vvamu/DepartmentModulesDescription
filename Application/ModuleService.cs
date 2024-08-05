using ConsoleApp1.Models;

namespace ConsoleApp1.Application;

internal class ModuleService
{
    
    public async Task Create()
    {
        var modules = ApplicationDbContext.SelectAll<Module>();
    }
  

}
