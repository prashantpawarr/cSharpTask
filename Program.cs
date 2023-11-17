using System;
using System.Data.SqlClient;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // Set EPPlus licensing context 
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Program obj = new Program();

        string name;
        int age;
        DateTime dob;
        string saveOption;
        string filePath;

        try
        {
            Console.Write("Enter your name: ");
            name = Console.ReadLine();

            Console.Write("Enter your age: ");
            age = Convert.ToInt32(Console.ReadLine());

            Console.Write("Enter your date of birth (MM-dd-yyyy): ");
            dob = DateTime.Parse(Console.ReadLine());

            Console.WriteLine("Name: " + name);
            Console.WriteLine("Age: " + age);
            Console.WriteLine("Date of Birth: " + dob.ToString("MM-dd-yyyy"));

            filePath = GetUserFilePath(); // Ask user for the file path

            Console.Write("Enter 'excel' to save to Excel, 'database' to save to Database, or 'notepad' to save to Notepad: ");
            saveOption = Console.ReadLine().ToLower();

            if (saveOption == "excel")
            {
                obj.SaveToExcel(name, age, dob, filePath);
                Console.WriteLine($"Data saved to Excel successfully! File Path: {filePath}");
            }
            else if (saveOption == "database")
            {
                obj.SaveToDatabase(name, age, dob);
                Console.WriteLine("Data saved to Database successfully!");
            }
            else if (saveOption == "notepad")
            {
                obj.SaveToNotepad(name, age, dob, filePath);
                Console.WriteLine($"Data saved to Notepad successfully! File Path: {filePath}");
            }
            else
            {
                Console.WriteLine("Invalid option. Data not saved.");
            }

        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // New method to prompt user for file path
    static string GetUserFilePath()
    {
        Console.Write("Enter the path of the file: ");
        return Console.ReadLine();
    }

    // Modify existing methods to accept file path as a parameter

    public void SaveToExcel(string name, int age, DateTime dob, string filePath)
    {
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("UserData");

            // This code is making Heading in Excel
            worksheet.Cells["A1"].Value = "Name";
            worksheet.Cells["B1"].Value = "Age";
            worksheet.Cells["C1"].Value = "Date of Birth";

            worksheet.Cells["A2"].Value = name;
            worksheet.Cells["B2"].Value = age;
            worksheet.Cells["C2"].Value = dob.ToString("MM-dd-yyyy");

            // Save the package to a file
            FileInfo excelFile = new FileInfo(filePath);
            excelPackage.SaveAs(excelFile);
        }
    }

    public void SaveToNotepad(string name, int age, DateTime dob, string filePath)
    {
        try
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                writer.WriteLine($"Name: {name}");
                writer.WriteLine($"Age: {age}");
                writer.WriteLine($"Date of Birth: {dob.ToString("MM-dd-yyyy")}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error saving to Notepad: " + ex.Message);
        }
    }
}
