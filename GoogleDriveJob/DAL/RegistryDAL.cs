using GoogleDriveJob.Models;
using Log4NetLibrary;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace GoogleDriveJob.DAL
{
    public class RegistryDAL
    {
        private string _connectionString;
        public RegistryDAL(IConfiguration iconfiguration)
        {
            _connectionString = iconfiguration.GetConnectionString("Default");
        }

        public List<TotalRevenueModel> GetTotalRevenueData(int clientId, DateTime from, DateTime to)
        {
            Logger.Info($"Get Total Revenue Data, ClientId: {clientId}, From: {from}, To: {to}");
            var totalRevenueModels = new List<TotalRevenueModel>();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    SqlCommand cmd = new SqlCommand("Report_GetTotalRevenue", con);
                    cmd.Parameters.Add(new SqlParameter("ClientId", clientId));
                    cmd.Parameters.Add(new SqlParameter("From", from));
                    cmd.Parameters.Add(new SqlParameter("To", to));
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        totalRevenueModels.Add(new TotalRevenueModel
                        {
                            OrderId = Convert.ToInt32(rdr[0]),
                            BuyerID = Convert.IsDBNull(rdr[1]) ? (int?)null : Convert.ToInt32(rdr[1]),
                            OrderTransaction = rdr[2].ToString(),
                            InsertDate = Convert.ToDateTime(rdr[3]).ToString("yyyyMMdd"),
                            SourceType = rdr[4].ToString(),
                            DeliveryOption = rdr[5].ToString(),
                            SKU = rdr[6].ToString(),
                            ProductPrice = Convert.ToDecimal(rdr[7]),
                            TotalRevenue = Convert.IsDBNull(rdr[6]) ? 0.00m : Convert.ToDecimal(rdr[8])
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return totalRevenueModels;
        }
        public List<NewUsersModel> GetNewUsersData(int clientId, DateTime from, DateTime to)
        {
            Logger.Info($"Get New Users Data, ClientId: {clientId}, From: {from}, To: {to}");
            var newUsersModels = new List<NewUsersModel>();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    SqlCommand cmd = new SqlCommand("Report_GetNewUsers", con);
                    cmd.Parameters.Add(new SqlParameter("ClientId", clientId));
                    cmd.Parameters.Add(new SqlParameter("From", from));
                    cmd.Parameters.Add(new SqlParameter("To", to));
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        newUsersModels.Add(new NewUsersModel
                        {
                            UserId = Convert.ToInt32(rdr[0]),
                            InsertDate = Convert.ToDateTime(rdr[1]).ToString("yyyyMMdd"),
                            SourceType = rdr[2].ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return newUsersModels;
        }
        public List<ListsCreatedModel> GetListsCreatedData(int clientId, DateTime from, DateTime to)
        {
            Logger.Info($"Get Lists Created Data, ClientId: {clientId}, From: {from}, To: {to}");
            var listsCreatedModels = new List<ListsCreatedModel>();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    SqlCommand cmd = new SqlCommand("Report_GetListsCreated", con);
                    cmd.Parameters.Add(new SqlParameter("ClientId", clientId));
                    cmd.Parameters.Add(new SqlParameter("From", from));
                    cmd.Parameters.Add(new SqlParameter("To", to));
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        listsCreatedModels.Add(new ListsCreatedModel
                        {
                            ListId = Convert.ToInt32(rdr[0]),
                            EventDate = Convert.ToDateTime(rdr[1]).ToString("yyyyMMdd"),
                            UniqueURL = rdr[2].ToString(),
                            Products = Convert.ToInt32(rdr[3]),
                            LastPurchasedDate = Convert.IsDBNull(rdr[4]) ? rdr[4].ToString() : Convert.ToDateTime(rdr[4]).ToString("yyyyMMdd"),
                            CreatedDate = Convert.ToDateTime(rdr[5]).ToString("yyyyMMdd"),
                            ListType = rdr[6].ToString()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listsCreatedModels;
        }
        public List<ListProductsModel> GetListProductsData(int clientId, DateTime from, DateTime to)
        {
            Logger.Info($"Get List Products Data, ClientId: {clientId}, From: {from}, To: {to}");
            var listProductsModels = new List<ListProductsModel>();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    SqlCommand cmd = new SqlCommand("Report_GetListProducts", con);
                    cmd.Parameters.Add(new SqlParameter("ClientId", clientId));
                    cmd.Parameters.Add(new SqlParameter("From", from));
                    cmd.Parameters.Add(new SqlParameter("To", to));
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        listProductsModels.Add(new ListProductsModel
                        {
                            InsertDate = Convert.ToDateTime(rdr[0]).ToString("yyyyMMdd"),
                            ListProductID = Convert.ToInt32(rdr[1]),
                            ListID = Convert.ToInt32(rdr[2]),
                            ProductID = Convert.ToInt32(rdr[3]),
                            ListProductStatus = rdr[4].ToString(),
                            Quantity = Convert.ToInt32(rdr[5]),
                            Price = Convert.ToDecimal(rdr[6])
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return listProductsModels;
        }
        public List<PopularProductsModel> GetPopularProductsData(int clientId, DateTime from, DateTime to)
        {
            Logger.Info($"Get Popular Products Data, ClientId: {clientId}, From: {from}, To: {to}");
            var popularProductsModels = new List<PopularProductsModel>();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    SqlCommand cmd = new SqlCommand("Report_GetPopularProducts", con);
                    cmd.Parameters.Add(new SqlParameter("ClientId", clientId));
                    cmd.Parameters.Add(new SqlParameter("From", from));
                    cmd.Parameters.Add(new SqlParameter("To", to));
                    cmd.CommandType = CommandType.StoredProcedure;
                    con.Open();
                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        popularProductsModels.Add(new PopularProductsModel
                        {
                            SKU = rdr[0].ToString(),
                            ProductTitle = rdr[1].ToString(),
                            Category = rdr[2].ToString(),
                            Price = Convert.ToDecimal(rdr[3]),
                            ItemsAdded = Convert.ToInt32(rdr[4]),
                            ItemsBought = Convert.ToInt32(rdr[5])
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return popularProductsModels;
        }

        public void CreateEmailLog(string email, string content, string subject)
        {
            Logger.Info($"Create Email Log, email: {email}, subject: {subject}");
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    String query = "INSERT INTO dbo.EmailLog (ClientId,Email,EmailContent,IsHTML,Subject,SenderName) VALUES (@clientId,@email,@emailContent,@isHtml, @subject,@senderName)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@clientId", 172);
                        command.Parameters.AddWithValue("@email", email);
                        command.Parameters.AddWithValue("@emailContent", content);
                        command.Parameters.AddWithValue("@isHtml", true);
                        command.Parameters.AddWithValue("@subject", subject);
                        command.Parameters.AddWithValue("@senderName", "GDS APP");

                        connection.Open();
                        int result = command.ExecuteNonQuery();

                        // Check Error
                        if (result < 0)
                            Logger.Error("Error inserting data into Database table EmailLog!");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex.Message);
            }
        }
    }
}
