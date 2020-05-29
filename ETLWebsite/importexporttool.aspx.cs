
using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class importexporttool : System.Web.UI.Page
{
    string constring = @"Data Source=.\SQLEXPRESS;Initial Catalog=etldb;Integrated Security=True";
    protected void Page_Load(object sender, EventArgs e)
    {
       
            StreamReader sr = new StreamReader(Server.MapPath("/files/filetype.txt"));
            string filetype = "";
            if (sr != null)
            {
                filetype = sr.ReadLine();
                sr.Close();
            }

            HyperLink1.NavigateUrl = "files/errorlog.txt";
            if (filetype != "")
                HyperLink2.NavigateUrl = "files/suppliertarget" + filetype;
        
    }

    protected void Button1_Click(object sender, EventArgs e)
    {
        string filename = FileUpload1.FileName.ToLower();
        string extension = filename.Substring(filename.LastIndexOf("."));
        string filetypepath = Server.MapPath("files/filetype.txt");
        if (File.Exists(filetypepath))
        {
            File.Delete(filetypepath);
        }
        StreamWriter sw = File.CreateText(filetypepath);
        sw.WriteLine(extension);
        sw.Flush();
        sw.Close();
        if (filename.EndsWith(".csv") || filename.EndsWith(".xls") || filename.EndsWith(".xlsx"))
        {
            string uploadfilename = "/files/suppliersource" + extension;
            FileUpload1.SaveAs(Server.MapPath(uploadfilename));
            Session["sourcefile"] = uploadfilename;
            importdatafromdatasource(uploadfilename, lblmsg1);
        }
        else
        {
            lblmsg1.Text = "Invalid file type";
        }
    }

    private void importdatafromdatasource(string uploadedfile, Label lbl)
    {
        try
        {
            string absolutepath = Server.MapPath(uploadedfile);
            if (absolutepath.ToLower().EndsWith(".csv"))
            {
                DataTable dt = GetDataTableFromCSVFile(absolutepath);
                Session["datatable"] = dt;
                lbl.Text = "Data Imported Successfully!!!";
                btnimporttodb.Enabled = true;

            }
            else
            {
                FileStream stream = File.Open(absolutepath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = null;
                if (absolutepath.ToLower().EndsWith(".xlsx"))
                {
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

                }
                else if (absolutepath.ToLower().EndsWith(".xls"))
                {
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                DataSet ds = excelReader.AsDataSet(conf);
                DataTable dt = new DataTable();
                dt = ds.Tables[0];
                Session["datatable"] = dt;
                lbl.Text = "Data Imported Successfully!!!";
                btnimporttodb.Enabled = true;
                //string s = "";
                //foreach (DataRow dr in dt.Rows)
                //{
                //    s += dr[0] + " " + dr[1] + " " + dr[2] + "<br/>";
                //}
                //lbl.Text = s;
            }
        }
        catch (Exception ex)
        {
            lbl.Text = ex.Message;
            btnimporttodb.Enabled = false;
        }
    }
    void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable FileData, string tablename)
    {
        using (SqlConnection dbConnection = new SqlConnection(constring))
        {
            dbConnection.Open();
            SqlCommand cmd = new SqlCommand("delete from " + tablename, dbConnection);
            cmd.ExecuteNonQuery();
            dbConnection.Close();
            dbConnection.Open();
            using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
            {
                s.DestinationTableName = tablename;
                foreach (var column in FileData.Columns)
                    s.ColumnMappings.Add(column.ToString(), column.ToString());
                s.WriteToServer(FileData);
            }
            dbConnection.Close();
        }
    }



    private static DataTable GetDataTableFromCSVFile(string csv_file_path)
    {
        DataTable csvData = new DataTable();
        try
        {
            using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
            {
                csvReader.SetDelimiters(new string[] { "," });
                csvReader.HasFieldsEnclosedInQuotes = true;
                string[] colFields = csvReader.ReadFields();
                foreach (string column in colFields)
                {
                    DataColumn datecolumn = new DataColumn(column);
                    datecolumn.AllowDBNull = true;
                    csvData.Columns.Add(datecolumn);
                }

                while (!csvReader.EndOfData)
                {
                    string[] fieldData = csvReader.ReadFields();
                    //Making empty value as null
                    for (int i = 0; i < fieldData.Length; i++)
                    {
                        if (fieldData[i] == "")
                        {
                            fieldData[i] = null;
                        }
                    }
                    csvData.Rows.Add(fieldData);
                }
            }
        }
        catch (Exception ex)
        {
        }
        return csvData;
    }

    protected void btnimporttodb_Click(object sender, EventArgs e)
    {
        try
        {
            Image1.Visible = true;
            DataTable FileDataTable = (DataTable)(Session["datatable"]);
            if (FileDataTable != null)
            {
                InsertDataIntoSQLServerUsingSQLBulkCopy(FileDataTable, "tblsuppliersource");
                lblmsg2.Text = "Data Imported To Database Successfully!!!";

            }
        }
        catch (Exception ex)
        {
            lblmsg2.Text = ex.Message;
        }
        Image1.Visible = false;
    }
    List<string> lstcode = new List<string>();
    //store source table data into destination table according to mapping template
    protected void btndestindata_Click(object sender, EventArgs e)
    {
        Image1.Visible = true;
        string Code;
        string Supplier_Name;
        string Country;
        string Contact_First_Name;
        string Contact_Last_Name;
        string Contact_Email;
        string Contact_Phone = null;
        string Postal_Address;
        string Notes;
        SqlConnection dbConnection = new SqlConnection(constring);
        dbConnection.Open();
        lstcode.Clear();
        SqlCommand cmdselect = new SqlCommand("select * from [tblsuppliersource]", dbConnection);
        SqlDataReader rdr = cmdselect.ExecuteReader();
        DataTable dt = new DataTable(); //data table to store mapping data
        dt.Columns.Add("Code");
        dt.Columns.Add("Supplier_Name");
        dt.Columns.Add("Country");
        dt.Columns.Add("Contact_First_Name");
        dt.Columns.Add("Contact_Last_Name");
        dt.Columns.Add("Contact_Email");
        dt.Columns.Add("Contact_Phone");
        dt.Columns.Add("Postal_Address");
        dt.Columns.Add("Notes");

        while (rdr.Read()) //read raw data from database table 
        {
            Supplier_Name = rdr["*ContactName"].ToString();
            String s = Regex.Replace(Supplier_Name, @"[^0-9A-Za-z ,]", "").Replace("&", "").Replace(" ", "");
            s = s.ToUpper();
            Code = abbrcode(s);
            Country = rdr["POCountry"].ToString();
            Contact_First_Name = rdr["FirstName"].ToString();
            Contact_Last_Name = rdr["LastName"].ToString();
            Contact_Email = rdr["EmailAddress"].ToString();
            if (rdr["MobileNumber"] != DBNull.Value)
                Contact_Phone = rdr["MobileNumber"].ToString();
            else
                Contact_Phone = rdr["PhoneNumber"].ToString();
            Postal_Address = rdr["POAddressLine2"] + " " + rdr["POAddressLine2"] + " " + rdr["POAddressLine3"] + " " + rdr["POAddressLine4"] + " " + rdr["POCity"] + " " + rdr["PORegion"] + " " + rdr["POPostalCode"];
            Notes = rdr["Website"].ToString();
            DataRow dtr = dt.NewRow();
            dtr[0] = Code;
            dtr[1] = Supplier_Name;
            dtr[2] = Country;
            dtr[3] = Contact_First_Name;
            dtr[4] = Contact_Last_Name;
            dtr[5] = Contact_Email;
            dtr[6] = Contact_Phone;
            dtr[7] = Postal_Address;
            dtr[8] = Notes;
            dt.Rows.Add(dtr);
        }
        dbConnection.Close();
        dbConnection.Open();
        SqlCommand deletecommand = new SqlCommand("delete from tblsuppliertarget", dbConnection);
        deletecommand.ExecuteNonQuery();
        dbConnection.Close();
        //store data into target database table
        string insertquery = "INSERT INTO tblsuppliertarget(Code,Supplier_Name,Country,Contact_First_Name,Contact_Last_Name,Contact_Email,Contact_Phone,Postal_Address,Notes) VALUES (@Code,@Supplier_Name,@Country,@Contact_First_Name,@Contact_Last_Name,@Contact_Email,@Contact_Phone,@Postal_Address,@Notes)";
        using (SqlCommand cmdinsert = new SqlCommand(insertquery, dbConnection))
        {
            dbConnection.Open();
            cmdinsert.Parameters.Add("@Code", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Supplier_Name", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Country", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Contact_First_Name", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Contact_Last_Name", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Contact_Email", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Contact_Phone", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Postal_Address", SqlDbType.VarChar);
            cmdinsert.Parameters.Add("@Notes", SqlDbType.VarChar);

            foreach (DataRow dr in dt.Rows)
            {
                Code = dr[0].ToString();
                Supplier_Name = dr[1].ToString();
                Country = dr[2].ToString();
                Contact_First_Name = dr[3].ToString();
                Contact_Last_Name = dr[4].ToString();
                Contact_Email = dr[5].ToString();
                Contact_Phone = dr[6].ToString();
                Postal_Address = dr[7].ToString();
                Notes = dr[8].ToString();

                cmdinsert.Parameters["@Code"].Value = Code;
                cmdinsert.Parameters["@Supplier_Name"].Value = Supplier_Name;
                cmdinsert.Parameters["@Country"].Value = Country;
                cmdinsert.Parameters["@Contact_First_Name"].Value = Contact_First_Name;
                cmdinsert.Parameters["@Contact_Last_Name"].Value = Contact_Last_Name;
                cmdinsert.Parameters["@Contact_Email"].Value = Contact_Email;
                cmdinsert.Parameters["@Contact_Phone"].Value = Contact_Phone;
                cmdinsert.Parameters["@Postal_Address"].Value = Postal_Address;
                cmdinsert.Parameters["@Notes"].Value = Notes;
                cmdinsert.ExecuteNonQuery();
            }
            lblmsg3.Text = "Data Stored To Target Successfully!!!";
            Image1.Visible = false;
        }
    }
    private string abbrcode(string code)
    {
        string r = "";
        if (code.Length >= 6)
        {
            r = code.Substring(0, 6);
        }
        else
        {
            r = code;
        }
        int count = 0;
        foreach (string s in lstcode)
        {
            if (s.Contains(r))
            {
                count++;
                //Console.Write("duplicate");
            }
        }
        if (count > 0)
        {
            r = r + count;
        }
        lstcode.Add(r);
        return r;
    }
    protected void btnExportData_Click(object sender, EventArgs e)
    {
        //create target file from target database
        StreamReader sr = new StreamReader(Server.MapPath("/files/filetype.txt"));
        string filetype = "";
        if (sr != null)
        {
            filetype = sr.ReadLine();
            sr.Close();
        }
        if (filetype.EndsWith("csv"))
        {
            using (SqlConnection con = new SqlConnection(constring))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT * FROM tblsuppliertarget"))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        cmd.Connection = con;
                        sda.SelectCommand = cmd;
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);

                            //Build the CSV file data as a Comma separated string.
                            string csv = string.Empty;

                            foreach (DataColumn column in dt.Columns)
                            {

                                //Add the Header row for CSV file.
                                csv += column.ColumnName + ',';
                            }

                            //Add new line.
                            csv += "\r\n";
                            int index = 0;
                            String path = Server.MapPath("files/errorlog.txt");
                            String targetfile = Server.MapPath("files/suppliertarget.csv");
                            if (File.Exists(path))
                            {
                                File.Delete(path);
                            }
                            StreamWriter sw = File.CreateText(path);
                            StreamWriter sw2 = File.CreateText(targetfile);
                            sw.AutoFlush = true;
                            sw2.AutoFlush = true;
                            String error = "";
                            bool haserrors = false;
                            foreach (DataRow row in dt.Rows)
                            {
                                index++;
                                error = validateData(row, index);
                                if (error != "")
                                {
                                    sw.WriteLine(error);
                                    haserrors = true;
                                }
                                error = "";
                                foreach (DataColumn column in dt.Columns)
                                {
                                    //Add the Data rows.
                                    csv += row[column.ColumnName].ToString().Replace(",", ";") + ',';
                                }

                                //Add new line.
                                csv += "\r\n";
                            }
                            sw.Flush();
                            sw.Close();
                            sw2.Write(csv);
                            sw2.Close();
                            HyperLink2.Visible = true;
                            if (haserrors)
                            {
                                lblmsg4.Text = "Data Exported with errors,<br/> click hyperlink to view error details";
                                Panel1.Visible = true;
                                HyperLink1.Visible = true;
                            }
                            else
                            {
                                lblmsg4.Text = "Data Exported Successfully!!!";
                                Panel1.Visible = false;
                                HyperLink1.Visible = false;
                            }
                        }
                    }
                }
            }
        }
        else
        {
            using (SqlConnection con = new SqlConnection(constring))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT * FROM tblsuppliertarget"))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter())
                    {
                        cmd.Connection = con;
                        sda.SelectCommand = cmd;
                        using (DataTable dt = new DataTable())
                        {
                            sda.Fill(dt);

                            try
                            {
                                //open file


                                String targetfile = Server.MapPath("files/suppliertarget" + filetype);

                                String path = Server.MapPath("files/errorlog.txt");

                                if (File.Exists(path))
                                {
                                    File.Delete(path);
                                }
                                StreamWriter sw = File.CreateText(path);
                                sw.AutoFlush = true;
                                String error = "";
                                bool haserrors = false;
                                                               int index = 0;
                                //write rows to excel file
                                for (int i = 0; i < (dt.Rows.Count); i++)
                                {
                                    index++;
                                    error = validateData(dt.Rows[i], index);
                                    if (error != "")
                                    {
                                        sw.WriteLine(error);
                                        haserrors = true;
                                    }
                                    error = "";
                                   
                                }
                               
                                sw.Flush();
                                sw.Close();
                                HyperLink2.Visible = true;
                                if (haserrors)
                                {
                                    lblmsg4.Text = "Data Exported with errors,<br/> click hyperlink to view error details";
                                    Panel1.Visible = true;
                                    HyperLink1.Visible = true;
                                }
                                else
                                {
                                    lblmsg4.Text = "Data Exported Successfully!!!";
                                    Panel1.Visible = false;
                                    HyperLink1.Visible = false;
                                }

                                ExportToExcel(dt, targetfile);

                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }
                }
            }
        }
    }

    public static void ExportToExcel(DataTable tbl, string excelFilePath = null)
    {

        //create a new ExcelPackage
        using (ExcelPackage excelPackage = new ExcelPackage())
        {
            //create a WorkSheet
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
            //add all the content from the DataTable, starting at cell A1
            worksheet.Cells["A1"].LoadFromDataTable(tbl, true);
            FileInfo fi = new FileInfo(excelFilePath);
            excelPackage.SaveAs(fi);
        }
    }
    private string validateData(DataRow dr, int index)
    {
        string error = "";
        if (dr[0].ToString() == "")
            error += "Code not found for row number :  " + index + "\n";

        if (dr[1].ToString() == "")
            error += "Supplier Name not found for row number : " + index + "\n";

        if (dr[2].ToString() == "")
            error += "Country not found for row number : " + index + "\n";

        //if (dr[5].ToString() == "")
        //    error += "Contact Email not found for row number " + index + "\n";
        //else if (!IsValid(dr[5].ToString()))
        //    error += "Contact Email is Invalid row number " + index + "\n";

        //if (dr[6].ToString() == "")
        //    error += "Contact Phone not found for row number " + index + "\n";

        return error;
    }
    public bool IsValid(string emailaddress)
    {
        try
        {
            MailAddress m = new MailAddress(emailaddress);
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
    }



    protected void Button2_Click(object sender, EventArgs e)
    {

        string filename = FileUpload2.FileName.ToLower();
        string extension = filename.Substring(filename.LastIndexOf("."));
        if (filename.EndsWith(".csv") || filename.EndsWith(".xls") || filename.EndsWith(".xlsx"))
        {
            string uploadfilename = "/files/suppliertarget" + extension;
            FileUpload2.SaveAs(Server.MapPath(uploadfilename));
            importdatafromdatasource(uploadfilename, lblmsg5);
            try
            {
                DataTable csvFileDataTable = (DataTable)(Session["datatable"]);
                if (csvFileDataTable != null)
                {
                    InsertDataIntoSQLServerUsingSQLBulkCopy(csvFileDataTable, "tblsuppliertarget");
                    lblmsg5.Text = "Data Imported To Database Successfully!!!";
                }
            }
            catch (Exception ex)
            {
                lblmsg5.Text = ex.Message;
            }
        }
        else
        {
            lblmsg5.Text = "Invalid file type";
        }
    }
}