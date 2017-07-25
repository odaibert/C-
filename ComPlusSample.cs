using System;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.EnterpriseServices;


namespace COMPlusServicesExample
{

    /// <summary>
    /// Summary description for OrderDB.
    /// </summary>
    [Transaction(TransactionOption.Supported)]
    public class OrderDB: ServicedComponent
    {
        //this class handles all data retrieval.
        public OrderDB()
        {
            
        }

        //method returns a connection object for the WideWorldImporters database.
        private SqlConnection CreateConnection()
        {
            //use Testing11@@ as SA user password
            string SQLIP = "172.22.34.185";
            SqlConnection databaseConnection = new SqlConnection("server=" + SQLIP + ";Trusted_Connection=false;uid=sa;pwd=Testing11@@;database=WideWorldImporters");
            return databaseConnection;
        } 

        //gets a dataset with the Orders
        public DataSet GetOrderDetails(int OrderId)
        {
            //grab the connection object
            SqlConnection currentConnection = CreateConnection();
            // Assumes that connection is a valid SqlConnection object.  

            //open the connection and create the command object
            currentConnection.Open();
            SqlCommand sqlCommandToExecute = new SqlCommand();

            //set the command type and provide the command text
            sqlCommandToExecute.CommandType = CommandType.Text;
            sqlCommandToExecute.CommandText = "Select * from Sales.OrderLines where OrderId = " + OrderId;
            sqlCommandToExecute.Connection = currentConnection;

            //create and fill the data adapter with the command object.
            SqlDataAdapter dataadapterOrder = new SqlDataAdapter(sqlCommandToExecute);
            
            //create the dataset
            DataSet datasetOrder = new DataSet("dsOrders");

            //fill the data set with the results in the data adapter
            dataadapterOrder.Fill(datasetOrder);

            //return the dataset
            return datasetOrder;
        }

       
        /// <summary>
        /// Summary Update Quantity for an Order.
        /// </summary>
        [Transaction(TransactionOption.Required)]
        public class Orders : ServicedComponent
        {

            //this class handles data modification.
            public Orders()
            {
            }

            //method returns a connection object for the Northwind database.
            private SqlConnection CreateConnection()
            {
                //use Testing11@@ as SA user password
                string SQLIP = "172.22.34.185";
                SqlConnection databaseConnection = new SqlConnection("server=" + SQLIP + ";Trusted_Connection=false;uid=sa;pwd=Testing11@@;database=WideWorldImporters");
                return databaseConnection;
            }

            //updates the detail quantity for the product.
            public bool UpdateOrderDetailQuantity(int OrderId, string StockItemId, int Quantity)
            {
                //grab the connection object
                SqlConnection currentConnection = CreateConnection();

                //open the connection object
                currentConnection.Open();

                try
                {
                    //try to update the quantity
                    //check to see if a transaction exists and if not the throw an exception.
                    if (!ContextUtil.IsInTransaction)
                        throw new Exception("This operation requires transaction");

                    //set the command text to be executed.
                    string CommandText = "UPDATE[Sales].[OrderLines] " +
                        "SET [Quantity] = " + Quantity + 
                        " WHERE OrderId = " + OrderId + 
                        " and StockItemID = " + StockItemId;
                     
                    //create the command object and set the connection property.
                    SqlCommand sqlCommandToExecute = new SqlCommand();
                    sqlCommandToExecute.Connection = currentConnection;

                    //set the command type and provide the command text.
                    sqlCommandToExecute.CommandType = CommandType.Text;
                    sqlCommandToExecute.CommandText = CommandText;
                    //execute the query.
                    sqlCommandToExecute.ExecuteNonQuery();

                    //complete the transaction.
                    ContextUtil.SetComplete();
                    return true;

                }

                catch (Exception e)
                {
                    //if an exception is thrown then close the connection object and abort the transaction.
                    currentConnection.Close();
                    ContextUtil.SetAbort();
                    return false;
                }
                finally
                {
                    //close the connection object.
                    currentConnection.Close();
                }
            }
        }

        public static implicit operator OrderDB(Orders v)
        {
            throw new NotImplementedException();
        }
    }
}
