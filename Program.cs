using System;
using System.Data.SqlClient;
using System.Data;
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
        string excelFilePath;

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

            Console.Write("Enter 'excel' to save to Excel or 'database' to save to Database: ");
            saveOption = Console.ReadLine().ToLower();

            if (saveOption == "excel")
            {
                excelFilePath = obj.SaveToExcel(name, age, dob);
                Console.WriteLine($"Data saved to Excel successfully! File Path: {excelFilePath}");
            }
            else if (saveOption == "database")
            {
                obj.SaveToDatabase(name, age, dob);
                Console.WriteLine("Data saved to Database successfully!");
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

    public string SaveToExcel(string name, int age, DateTime dob)
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
            FileInfo excelFile = new FileInfo("UserData.xlsx");
            excelPackage.SaveAs(excelFile);

            return excelFile.FullName;
        }
    }

    public void SaveToDatabase(string name, int age, DateTime dob)
    {
        string connectionString = "Data Source=CTS-D049;Initial Catalog=crudApplication;Integrated Security=True";
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            string insertQuery = "INSERT INTO UserDetails (Name, Age, DateOfBirth) VALUES (@Name, @Age, @DateOfBirth)";
            using (SqlCommand command = new SqlCommand(insertQuery, connection))
            {
                command.Parameters.AddWithValue("@Name", name);
                command.Parameters.AddWithValue("@Age", age);
                command.Parameters.AddWithValue("@DateOfBirth", dob);

                command.ExecuteNonQuery();
            }
        }
    }
}

