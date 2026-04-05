using LanguageCourseSlides.Forms;

namespace LanguageCourseSlides;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}
