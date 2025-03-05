using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Windows.Forms;

public class ModelValidator
{
    private readonly string _connectionString;

    public ModelValidator()
    {
        // Fetch connection string from app.config
        _connectionString = PaperlessQc.Properties.Settings.Default.BkConString;
    }

    /// <summary>
    /// Checks if the model exists in the TB_M_ITEM table.
    /// </summary>
    /// <param name="modelId">The model ID to validate.</param>
    /// <returns>True if the model exists, otherwise false.</returns>
    public bool IsModelValid(string modelId)
    {
        bool exists = false;
        string query = "SELECT COUNT(*) FROM TB_M_ITEM WHERE ITEM_ID = @ModelId";

        try
        {
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@ModelId", modelId);
                    int count = (int)cmd.ExecuteScalar();
                    exists = count > 0;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        return exists;
    }
}
