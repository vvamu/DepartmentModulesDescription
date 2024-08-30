namespace ConsoleApp1.Helpers;

public static class TranslationHelper
{
    public static async Task<string> TranslateToRu(string str = " ", GTranslatorAPI.Languages sourceLanguage = GTranslatorAPI.Languages.be, GTranslatorAPI.Languages targetLanguage = GTranslatorAPI.Languages.ru)
    {
        var translator = new GTranslatorAPI.GTranslatorAPIClient();
        try
        {

            //var result = await translator.TranslateAsync(GTranslatorAPI.Languages.be, GTranslatorAPI.Languages.ru, "Добрага ранку");
            var result = await translator.TranslateAsync(GTranslatorAPI.Languages.be, GTranslatorAPI.Languages.ru, str.Replace("«", "\"").Replace("»", "\""));

            return result.TranslatedText;
        }
        catch (Exception ex) { }
        return "";
    }

    public static async Task<ConsoleApp1.Models.Module> TranslateToRu(ConsoleApp1.Models.Module module)
    {

        module.Description = await TranslateToRu(module.Description);
        module.Name = await TranslateToRu(module.Name);
        module.Speciality = await TranslateToRu(module.Speciality);
        return module;
    }
}
