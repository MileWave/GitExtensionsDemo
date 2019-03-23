using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BackEnd
{
    public class Operations
    {

        OleDbConnectionStringBuilder Builder = new OleDbConnectionStringBuilder
        {
            Provider = "Microsoft.ACE.OLEDB.12.0",
            DataSource = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CustomersDatabase.accdb")
        };

        Exception mExceptiom;
        public Exception Exception
        {
            get
            {
                return mExceptiom;
            }
        }
        /// <summary>
        /// Read all customer into a DataTable
        /// </summary>
        /// <returns></returns>
        public DataTable ReadCustomers()
        {
            var dt = new DataTable();

            try
            {
                using (OleDbConnection cn = new OleDbConnection { ConnectionString = Builder.ConnectionString })
                {
                    using (OleDbCommand cmd = new OleDbCommand { Connection = cn })
                    {

                        cmd.CommandText = "SELECT Identifier,CompanyName, ContactName, ContactTitle, [Address], City,PostalCode, Country FROM  Customers";

                        cn.Open();

                        dt.Load(cmd.ExecuteReader());
                        dt.Columns["Identifier"].ColumnMapping = MappingType.Hidden;
                    }
                }
            }
            catch (Exception ex)
            {
                mExceptiom = ex;
            }

            return dt;

        }
        /// <summary>
        /// Remove a single customer by primary key
        /// </summary>
        /// <param name="pIdentifier"></param>
        /// <returns></returns>
        public bool RemoveCustomer(int pIdentifier)
        {
            bool Success = true;

            try
            {
                using (OleDbConnection cn = new OleDbConnection { ConnectionString = Builder.ConnectionString })
                {
                    using (OleDbCommand cmd = new OleDbCommand { Connection = cn })
                    {

                        cmd.CommandText = "DELETE FROM Customers WHERE Identifier = @Identifier";

                        cmd.Parameters.AddWithValue("@Identifier", pIdentifier);

                        cn.Open();

                        int Affected = cmd.ExecuteNonQuery();

                        if (Affected == 1)
                        {
                            Success = true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            catch (OleDbException oex)
            {
                Success = false;
                mExceptiom = oex;
            }
            catch (Exception ex)
            {
                Success = false;
                mExceptiom = ex;
            }

            return Success;

        }
        /// <summary>
        /// Update a single customer by passing a DataRow, we use the primary
        /// key to locate the record and do the update using existing field values
        /// in the DataRow passed in. Others might have a parameter for each field 
        /// rather than passing a data row or perhaps an instance of a concrete class.
        /// </summary>
        /// <param name="pRow"></param>
        /// <returns></returns>
        public bool UpdateRow(DataRow pRow)
        {
            bool Success = true;

            try
            {
                using (OleDbConnection cn = new OleDbConnection { ConnectionString = Builder.ConnectionString })
                {
                    using (OleDbCommand cmd = new OleDbCommand { Connection = cn })
                    {

                        cmd.CommandText = "UPDATE Customers SET CompanyName = @CompanyName, ContactName = @ContactName, ContactTitle = @ContactTitle, " +
                                          "[Address] = @Address, City = @City,PostalCode = @PostalCode, Country = @Country  WHERE Identifier = @Identifier";


                        cmd.Parameters.AddWithValue("@CompanyName", pRow.Field<String>("CompanyName"));
                        cmd.Parameters.AddWithValue("@ContactName", pRow.Field<String>("ContactName"));
                        cmd.Parameters.AddWithValue("@ContactTitle", pRow.Field<String>("ContactTitle"));
                        cmd.Parameters.AddWithValue("@Address", pRow.Field<String>("Address"));
                        cmd.Parameters.AddWithValue("@City", pRow.Field<String>("City"));
                        cmd.Parameters.AddWithValue("@PostalCode", pRow.Field<String>("PostalCode"));
                        cmd.Parameters.AddWithValue("@Country", pRow.Field<String>("Country"));
                        cmd.Parameters.AddWithValue("@Identifier", pRow.Field<int>("Identifier"));

                        cn.Open();

                        int Affected = cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (OleDbException oex)
            {
                Success = false;
                mExceptiom = oex;
            }
            catch (Exception ex)
            {
                Success = false;
                mExceptiom = ex;
            }

            return Success;
        }
        /// <summary>
        /// Add a new record using a DataRow.
        /// pIdentifier will be set if the INSERT was successful, otherwise in the
        /// caller it will be 0.
        /// </summary>
        /// <param name="pRow"></param>
        /// <param name="pIdentifier"></param>
        /// <returns></returns>
        public bool AddNewRow(DataRow pRow, ref int pIdentifier)
        {
            bool Success = true;

            try
            {
                using (OleDbConnection cn = new OleDbConnection { ConnectionString = Builder.ConnectionString })
                {
                    using (OleDbCommand cmd = new OleDbCommand { Connection = cn })
                    {

                        cmd.CommandText = "INSERT INTO Customers (CompanyName,ContactName,ContactTitle,[Address],City,PostalCode,Country) " +
                                          "VALUES (@CompanyName,@ContactName,@ContactTitle,@Address,@City,@PostalCode,@Country)";


                        cmd.Parameters.AddWithValue("@CompanyName", pRow.Field<String>("CompanyName"));
                        cmd.Parameters.AddWithValue("@ContactName", pRow.Field<String>("ContactName"));
                        cmd.Parameters.AddWithValue("@ContactTitle", pRow.Field<String>("ContactTitle"));
                        cmd.Parameters.AddWithValue("@Address", pRow.Field<String>("Address"));
                        cmd.Parameters.AddWithValue("@City", pRow.Field<String>("City"));
                        cmd.Parameters.AddWithValue("@PostalCode", pRow.Field<String>("PostalCode"));
                        cmd.Parameters.AddWithValue("@Country", pRow.Field<String>("Country"));

                        cn.Open();

                        int Affected = cmd.ExecuteNonQuery();

                        if (Affected == 1)
                        {
                            cmd.CommandText = "Select @@Identity";
                            pIdentifier = Convert.ToInt32(cmd.ExecuteScalar());
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            catch (OleDbException oex)
            {
                Success = false;
                mExceptiom = oex;
            }
            catch (Exception ex)
            {
                Success = false;
                mExceptiom = ex;
            }

            return Success;
        }

    }
}
