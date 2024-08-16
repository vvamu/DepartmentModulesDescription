namespace ConsoleApp1.Helpers;

public static class TranslationHelper
{
    public static async Task<string> TranslateToRu(string str, GTranslatorAPI.Languages sourceLanguage = GTranslatorAPI.Languages.be, GTranslatorAPI.Languages targetLanguage = GTranslatorAPI.Languages.ru)
    {
        var translator = new GTranslatorAPI.GTranslatorAPIClient();
        var result = await translator.TranslateAsync(sourceLanguage, targetLanguage, str);
        return result.TranslatedText;
    }

    public static async Task TranslateToRu(ConsoleApp1.Models.Module module)
    {
        module.Description = await TranslateToRu(module.Description);
        module.Name = await TranslateToRu(module.Name);
        module.Speciality = await TranslateToRu(module.Speciality);
    }
}
