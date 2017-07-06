protected void ExportMasterCard(object sender, EventArgs e)
{
    string connStr = ConfigurationManager.ConnectionString["DBCS"].ConnectionString;
    using(SqlConnection con = new SqlConnection(connStr))
    {
        using(SqlCommand cmd = new SqlCommand("Select * from MasterCard"))
        {
            using(SqlDataAdapter sda = new SqlDataAdapter())
            {
                cmd.Connection = con;
                sda.SelectCommand = cmd;
                using(DataTable dt = new DataTable())
                {
                    sda.Fill(dt);

                    string csv = string.Empty;

                    foreach (DataColumn column in dt.Columns)
                    {
                        csv+= column.ColumnName+',';
                    }
                    csv += "\r\n";
                    foreach(DataRow row in dr.Rows)
                    {
                        foreach(DataColumn column in dt.Columns)
                        {
                            csv += row[column.ColumnName].ToString().Replace(",",";")+',';
                        }
                        csv += "\r\n";
                    }
                    Response.Clear();
                    Response.Buffer = true;
                    Response.AddHeader("content-disposition","attachment;filename=MasterCardFile.csv");
                    Response.Charset = "";
                    Response.ContentType = "application/text";
                    Response.Output.Write(csv);
                    Response.Flush();
                    Response.End();
                }
            }
        }
    }
}