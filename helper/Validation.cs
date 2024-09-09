using System.Dynamic;

public static class Validation
{

    public static bool FolderExistsValidation()
    {
        var inputFolder = Directory.GetDirectories(Path.Combine(Directory.GetCurrentDirectory(), "data", "input"));

        if (inputFolder.Length < 1) return false;
        return true;
    }
}