using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlServerCe;
using System.Data;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace Coil_Diagnostor.Function
{
    public class FunctionDataControl
    {
        // 데이터베이스 연결
        public string connectionString = @"Data Source=|DataDirectory|\CoilDiagnostorDB.sdf;pwd=realgain";

        public SqlCeConnection conn = null;
        public SqlCeCommand cmd = null;
        public SqlCeTransaction tran = null;

        #region ▣ 환경 설정

        #region ▣ 환경 설정 - 기준값(RCS)

        public DataTable GetRCSReferenceValueOhDegreeInfo(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT PlantName, Hogi, Oh_Degree ";
                _select = _select + "FROM tbRCSReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetRCSReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("RdcStop_RCSReferenceValue");
            dt.Columns.Add("RdcUp_RCSReferenceValue");
            dt.Columns.Add("RdcMove_RCSReferenceValue");
            dt.Columns.Add("RacStop_RCSReferenceValue");
            dt.Columns.Add("RacUp_RCSReferenceValue");
            dt.Columns.Add("RacMove_RCSReferenceValue");
            dt.Columns.Add("LStop_RCSReferenceValue");
            dt.Columns.Add("LUp_RCSReferenceValue");
            dt.Columns.Add("LMove_RCSReferenceValue");
            dt.Columns.Add("CStop_RCSReferenceValue");
            dt.Columns.Add("CUp_RCSReferenceValue");
            dt.Columns.Add("CMove_RCSReferenceValue");
            dt.Columns.Add("QStop_RCSReferenceValue");
            dt.Columns.Add("QUp_RCSReferenceValue");
            dt.Columns.Add("QMove_RCSReferenceValue");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree";
                _select = _select + ", RdcStop_RCSReferenceValue, RdcUp_RCSReferenceValue, RdcMove_RCSReferenceValue";
                _select = _select + ", RacStop_RCSReferenceValue, RacUp_RCSReferenceValue, RacMove_RCSReferenceValue";
                _select = _select + ", LStop_RCSReferenceValue, LUp_RCSReferenceValue, LMove_RCSReferenceValue";
                _select = _select + ", CStop_RCSReferenceValue, CUp_RCSReferenceValue, CMove_RCSReferenceValue";
                _select = _select + ", QStop_RCSReferenceValue, QUp_RCSReferenceValue, QMove_RCSReferenceValue ";
                _select = _select + "FROM tbRCSReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public int GetRCSReferenceValueDataCount(string _strPlantName, string _strHogi, string _strOhDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbRCSReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public int GetRCSReferenceValueDataMaxOhDegree(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            int iOh_Degree = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT MAX(Oh_Degree) AS Oh_Degree ";
                _select = _select + "FROM tbRCSReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iOh_Degree = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["Oh_Degree"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iOh_Degree;
        }

        public bool InsertRCSReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree, string[] _arrayData)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbRCSReferenceValue(PlantName, Hogi, Oh_Degree, RdcStop_RCSReferenceValue, RdcUp_RCSReferenceValue";
                cmd.CommandText = cmd.CommandText + ", RdcMove_RCSReferenceValue, RacStop_RCSReferenceValue, RacUp_RCSReferenceValue, RacMove_RCSReferenceValue";
                cmd.CommandText = cmd.CommandText + ", LStop_RCSReferenceValue, LUp_RCSReferenceValue, LMove_RCSReferenceValue, CStop_RCSReferenceValue";
                cmd.CommandText = cmd.CommandText + ", CUp_RCSReferenceValue, CMove_RCSReferenceValue, QStop_RCSReferenceValue, QUp_RCSReferenceValue";
                cmd.CommandText = cmd.CommandText + ", QMove_RCSReferenceValue) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOhDegree.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[0].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[1].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[2].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[3].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[4].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[5].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[6].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[7].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[8].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[9].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[10].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[11].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[12].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[13].ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _arrayData[14].ToString().Trim() + "') ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool UpDateRCSReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree, string[] _arrayData)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbRCSReferenceValue SET ";
                cmd.CommandText = cmd.CommandText + "RdcStop_RCSReferenceValue = '" + _arrayData[0].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", RdcUp_RCSReferenceValue = '" + _arrayData[1].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", RdcMove_RCSReferenceValue = '" + _arrayData[2].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", RacStop_RCSReferenceValue = '" + _arrayData[3].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", RacUp_RCSReferenceValue = '" + _arrayData[4].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", RacMove_RCSReferenceValue = '" + _arrayData[5].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", LStop_RCSReferenceValue = '" + _arrayData[6].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", LUp_RCSReferenceValue = '" + _arrayData[7].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", LMove_RCSReferenceValue = '" + _arrayData[8].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", CStop_RCSReferenceValue = '" + _arrayData[9].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", CUp_RCSReferenceValue = '" + _arrayData[10].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", CMove_RCSReferenceValue = '" + _arrayData[11].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", QStop_RCSReferenceValue = '" + _arrayData[12].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", QUp_RCSReferenceValue = '" + _arrayData[13].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", QMove_RCSReferenceValue = '" + _arrayData[14].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOhDegree.ToString().Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #region ▣ 환경 설정 - 기준값(DRPI)

        public DataTable GetDRPIReferenceValueOhDegreeInfo(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT PlantName, Hogi, Oh_Degree ";
                _select = _select + "FROM tbDRPIReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("CoilCabinetType");
            dt.Columns.Add("CoilCabinetName");
            dt.Columns.Add("Rdc_DRPIReferenceValue");
            dt.Columns.Add("Rac_DRPIReferenceValue");
            dt.Columns.Add("L_DRPIReferenceValue");
            dt.Columns.Add("C_DRPIReferenceValue");
            dt.Columns.Add("Q_DRPIReferenceValue");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree, CoilCabinetType, CoilCabinetName";
                _select = _select + ", Rdc_DRPIReferenceValue, Rac_DRPIReferenceValue, L_DRPIReferenceValue, C_DRPIReferenceValue, Q_DRPIReferenceValue ";
                _select = _select + "FROM tbDRPIReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, CoilCabinetType,CoilCabinetName ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree
            , string _strCoilCabinetType, string _strCoilCabinetName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("CoilCabinetType");
            dt.Columns.Add("CoilCabinetName");
            dt.Columns.Add("Rdc_DRPIReferenceValue");
            dt.Columns.Add("Rac_DRPIReferenceValue");
            dt.Columns.Add("L_DRPIReferenceValue");
            dt.Columns.Add("C_DRPIReferenceValue");
            dt.Columns.Add("Q_DRPIReferenceValue");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree, CoilCabinetType, CoilCabinetName";
                _select = _select + ", Rdc_DRPIReferenceValue, Rac_DRPIReferenceValue, L_DRPIReferenceValue, C_DRPIReferenceValue, Q_DRPIReferenceValue ";
                _select = _select + "FROM tbDRPIReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                _select = _select + "AND CoilCabinetType = '" + _strCoilCabinetType.Trim() + "' ";
                _select = _select + "AND CoilCabinetName = '" + _strCoilCabinetName.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, CoilCabinetType, CoilCabinetName ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public int GetDRPIReferenceValueDataCount(string _strPlantName, string _strHogi, string _strOhDegree
            , string _strCoilCabinetType, string _strCoilCabinetName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbDRPIReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                _select = _select + "AND CoilCabinetType = '" + _strCoilCabinetType.Trim() + "' ";
                _select = _select + "AND CoilCabinetName = '" + _strCoilCabinetName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

		public int GetDRPIReferenceValueDataCount(string _strPlantName, string _strHogi, string _strOhDegree)
		{
			//----- Sql Open
			conn = new SqlCeConnection(connectionString);
			conn.Open();

			DataTable dt = new DataTable();

			dt.Columns.Add("iCount");

			int iCount = 0;

			try
			{
				// 커맨드에 커넥션을 연결
				string _select = "";
				_select = "SELECT Count(*) AS iCount ";
				_select = _select + "FROM tbDRPIReferenceValue ";
				_select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
				_select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
				_select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";

				cmd = new SqlCeCommand(_select, conn);
				SqlCeDataReader rdr = cmd.ExecuteReader();

				// DB에서 가져온 Data를 DataSet에 바인딩
				dt.Load(rdr);

				iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
			}

			conn.Close();
			return iCount;
		}

        public int GetDRPIReferenceValueDataMaxOhDegree(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            int iOh_Degree = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT MAX(Oh_Degree) AS Oh_Degree ";
                _select = _select + "FROM tbDRPIReferenceValue ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iOh_Degree = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["Oh_Degree"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iOh_Degree;
        }

        public bool InsertDRPIReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree
            , string _strCoilCabinetType, string _strCoilCabinetName, string[] _arrayData)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;


                cmd.CommandText = "INSERT INTO tbDRPIReferenceValue(PlantName, Hogi, Oh_Degree, CoilCabinetType, CoilCabinetName";
                cmd.CommandText = cmd.CommandText + ", Rdc_DRPIReferenceValue, Rac_DRPIReferenceValue, L_DRPIReferenceValue";
                cmd.CommandText = cmd.CommandText + ", C_DRPIReferenceValue, Q_DRPIReferenceValue) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOhDegree.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strCoilCabinetType.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strCoilCabinetName.ToString().Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _arrayData[0].ToString().Trim() + ", ";
                cmd.CommandText = cmd.CommandText + _arrayData[1].ToString().Trim() + ", ";
                cmd.CommandText = cmd.CommandText + _arrayData[2].ToString().Trim() + ", ";
                cmd.CommandText = cmd.CommandText + _arrayData[3].ToString().Trim() + ", ";
                cmd.CommandText = cmd.CommandText + _arrayData[4].ToString().Trim() + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool UpDateDRPIReferenceValueDataInfo(string _strPlantName, string _strHogi, string _strOhDegree
            , string _strCoilCabinetType, string _strCoilCabinetName, string[] _arrayData)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbDRPIReferenceValue SET ";
                cmd.CommandText = cmd.CommandText + "Rdc_DRPIReferenceValue = '" + _arrayData[0].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", Rac_DRPIReferenceValue = '" + _arrayData[1].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", L_DRPIReferenceValue   = '" + _arrayData[2].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", C_DRPIReferenceValue   = '" + _arrayData[3].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + ", Q_DRPIReferenceValue   = '" + _arrayData[4].Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOhDegree.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilCabinetType = '" + _strCoilCabinetType.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilCabinetName = '" + _strCoilCabinetName.ToString().Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #endregion


        #region ▣ 영점 보정

        #region ▣ 영점 보정 - 일반 모드 / 휘스톤 모드

        public DataTable GetSetOffsetDataGridViewDataInfo(string _strPlantName, string _strMeasurementMode, string _strDAMRelayName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("DAMRelayName");
            dt.Columns.Add("ChannelName");
            dt.Columns.Add("DC_Resistance");
            dt.Columns.Add("AC_Resistance");
            dt.Columns.Add("Inductance");
            dt.Columns.Add("Capacitance");
            dt.Columns.Add("Sorting");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, DAMRelayName, MeasurementMode, ChannelName";
                _select = _select + ", DC_Resistance, AC_Resistance, Inductance, Capacitance, Sorting ";
                _select = _select + "FROM tbSetOffset ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND MeasurementMode = '" + _strMeasurementMode.Trim() + "' ";
                _select = _select + "AND DAMRelayName = '" + _strDAMRelayName.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, DAMRelayName, Sorting, ChannelName ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetSetOffsetDataGridViewDataInfo(string _strPlantName, string _strMeasurementMode, string _strDAMRelayName, string _strChannelName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("DAMRelayName");
            dt.Columns.Add("ChannelName");
            dt.Columns.Add("DC_Resistance");
            dt.Columns.Add("AC_Resistance");
            dt.Columns.Add("Inductance");
            dt.Columns.Add("Capacitance");
            dt.Columns.Add("Sorting");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, DAMRelayName, MeasurementMode, ChannelName";
                _select = _select + ", DC_Resistance, AC_Resistance, Inductance, Capacitance, Sorting ";
                _select = _select + "FROM tbSetOffset ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND MeasurementMode = '" + _strMeasurementMode.Trim() + "' ";
                _select = _select + "AND DAMRelayName = '" + _strDAMRelayName.Trim() + "' ";
                _select = _select + "AND ChannelName = '" + _strChannelName.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, DAMRelayName, Sorting, ChannelName ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public int GetSetOffsetDataGridViewDataCount(string _strPlantName, string _strMeasurementMode, string _strDAMRelayName, string _strChannelName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbSetOffset ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND MeasurementMode = '" + _strMeasurementMode.Trim() + "' ";
                _select = _select + "AND DAMRelayName = '" + _strDAMRelayName.Trim() + "' ";
                _select = _select + "AND ChannelName = '" + _strChannelName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertSetOffsetDataGridViewDataInfo(string _strPlantName, string _strMeasurementMode, string _strDAMRelayName, string _strChannelName
            , decimal _dcmDC_Resistance, decimal _dcmAC_Resistance, decimal _dcmInductance, decimal _dcmCapacitance)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbSetOffset(PlantName, MeasurementMode, DAMRelayName, ChannelName";
                cmd.CommandText = cmd.CommandText + ", DC_Resistance, AC_Resistance, Inductance, Capacitance) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementMode.ToString().Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _strDAMRelayName.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strChannelName.ToString().Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _dcmDC_Resistance + ", ";
                cmd.CommandText = cmd.CommandText + _dcmAC_Resistance + ", ";
                cmd.CommandText = cmd.CommandText + _dcmInductance + ", ";
                cmd.CommandText = cmd.CommandText + _dcmCapacitance + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool UpDateSetOffsetDataGridViewDataInfo(string _strPlantName, string _strMeasurementMode, string _strDAMRelayName, string _strChannelName
            , decimal _dcmDC_Resistance, decimal _dcmAC_Resistance, decimal _dcmInductance, decimal _dcmCapacitance)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbSetOffset SET ";
                cmd.CommandText = cmd.CommandText + "DC_Resistance = " + _dcmDC_Resistance + " ";
                cmd.CommandText = cmd.CommandText + ", AC_Resistance = " + _dcmAC_Resistance + " ";
                cmd.CommandText = cmd.CommandText + ", Inductance = " + _dcmInductance + " ";
                cmd.CommandText = cmd.CommandText + ", Capacitance = " + _dcmCapacitance + " ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND MeasurementMode = '" + _strMeasurementMode.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DAMRelayName = '" + _strDAMRelayName.ToString().Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ChannelName = '" + _strChannelName.ToString().Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #endregion


        #region ▣ RCS 측정

        public DataTable GetDRPIReferenceValueOhDegreeData(string _strPlantName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Hogi");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT Hogi ";
                _select = _select + "FROM tbDRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "ORDER BY Hogi ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetRCSReferenceValueOhDegreeData(string _strPlantName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Hogi");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT Hogi ";
                _select = _select + "FROM tbRCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "ORDER BY Hogi ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        #region ▣ RCS 측정 - Header

        public string GetRCSDiagnosisHogiMaxOhDegree(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            string strOh_Degree = "";

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Max(Oh_Degree) AS Oh_Degree ";
                _select = _select + "FROM tbRCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                strOh_Degree = dt == null || dt.Rows.Count < 1 ? "제 0 차" : dt.Rows[0]["Oh_Degree"].ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return strOh_Degree;
        }

        public int GetRCSDiagnosisHeaderDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbRCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi
            , string _strOh_Degree, string _strPowerCabinet, string _strControlName, decimal _dcmTemperature_ReferenceValue
            , decimal _dcmTemperatureUpDown_ReferenceValue, string _strFrequency, string _strVoltageLevel
            , int _intMeasurementCount, int _intItem_Rdc, int _intItem_Rac
            , int _intItem_L, int _intItem_C, int _intItem_Q, string _strMeasurementMode
            , string _strMeasurementDate, string _strMeasurementResult)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbRCSDiagnosisHeader(PlantName, Hogi, Oh_Degree, PowerCabinet, ControlName";
                cmd.CommandText = cmd.CommandText + ", Temperature_ReferenceValue, TemperatureUpDown_ReferenceValue, Frequency, VoltageLevel";
                cmd.CommandText = cmd.CommandText + ", MeasurementCount, Item_Rdc, Item_Rac, Item_L, Item_C, Item_Q";
                cmd.CommandText = cmd.CommandText + ", MeasurementMode, MeasurementDate, MeasurementResult) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strPowerCabinet.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _dcmTemperature_ReferenceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmTemperatureUpDown_ReferenceValue + ", '";
                cmd.CommandText = cmd.CommandText + _strFrequency.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strVoltageLevel.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intMeasurementCount + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_L + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_C + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Q + ", '";
                cmd.CommandText = cmd.CommandText + _strMeasurementMode.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementResult.Trim() + "') ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool InsertRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi
            , string _strOh_Degree, string _strPowerCabinet, string _strControlName, decimal _dcmTemperature_ReferenceValue
            , decimal _dcmTemperatureUpDown_ReferenceValue, string _strFrequency, string _strVoltageLevel
            , int _intMeasurementCount, int _intItem_Rdc, int _intItem_Rac
            , int _intItem_L, int _intItem_C, int _intItem_Q, string _strMeasurementMode
            , string _strMeasurementDate, string _strMeasurementResult, int _intTAKE_ID)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbRCSDiagnosisHeader(PlantName, Hogi, Oh_Degree, PowerCabinet, ControlName";
                cmd.CommandText = cmd.CommandText + ", Temperature_ReferenceValue, TemperatureUpDown_ReferenceValue, Frequency, VoltageLevel";
                cmd.CommandText = cmd.CommandText + ", MeasurementCount, Item_Rdc, Item_Rac, Item_L, Item_C, Item_Q";
                cmd.CommandText = cmd.CommandText + ", MeasurementMode, MeasurementDate, MeasurementResult, TAKE_ID) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strPowerCabinet.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _dcmTemperature_ReferenceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmTemperatureUpDown_ReferenceValue + ", '";
                cmd.CommandText = cmd.CommandText + _strFrequency.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strVoltageLevel.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intMeasurementCount + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_L + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_C + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Q + ", '";
                cmd.CommandText = cmd.CommandText + _strMeasurementMode.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementResult.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intTAKE_ID + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi
            , string _strOh_Degree, string _strPowerCabinet, string _strControlName)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbRCSDiagnosisHeader ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #region ▣ RCS 측정 - Detail

        public int GetRCSDiagnosisDetailDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlName, string _intCoilName, int _intCoilNumber)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbRCSDiagnosisDetail ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";
                _select = _select + "AND CoilName = '" + _intCoilName.Trim() + "' ";
                _select = _select + "AND CoilNumber = " + _intCoilNumber + " ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public int GetRCSDiagnosisDetailDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbRCSDiagnosisDetail ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertRCSDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlName, string _strCoilName, int _intCoilNumber, decimal _dcmDC_ResistanceValue
            , decimal _dcmDC_Deviation, decimal _dcmAC_ResistanceValue, decimal _dcmAC_Deviation, decimal _dcmL_InductanceValue
            , decimal _dcmL_Deviation, decimal _dcmC_CapacitanceValue, decimal _dcmC_Deviation, decimal _dcmQ_FactorValue
			, decimal _dcmQ_Deviation, string _strResult, string _strMeasurementDate, int _intSeqNumber, int _intControlSeqNumber)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbRCSDiagnosisDetail(PlantName, Hogi, Oh_Degree, PowerCabinet, ControlName, CoilName, CoilNumber";
                cmd.CommandText = cmd.CommandText + ", DC_ResistanceValue, DC_Deviation, AC_ResistanceValue, AC_Deviation";
                cmd.CommandText = cmd.CommandText + ", L_InductanceValue, L_Deviation, C_CapacitanceValue, C_Deviation, Q_FactorValue, Q_Deviation";
				cmd.CommandText = cmd.CommandText + ", Result, MeasurementDate, SeqNumber, ControlSeqNumber) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strPowerCabinet.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strCoilName.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intCoilNumber + ", ";
                cmd.CommandText = cmd.CommandText + _dcmDC_ResistanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmDC_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmAC_ResistanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmAC_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmL_InductanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmL_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmC_CapacitanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmC_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmQ_FactorValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmQ_Deviation + ", '";
                cmd.CommandText = cmd.CommandText + _strResult + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementDate.Trim() + "', ";
				cmd.CommandText = cmd.CommandText + _intSeqNumber + ", ";
				cmd.CommandText = cmd.CommandText + _intControlSeqNumber + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteRCSDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree, string _strPowerCabinet)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbRCSDiagnosisDetail ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteRCSDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlName, string _strCoilName, int _intCoilNumber)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbRCSDiagnosisDetail ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strCoilName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilNumber = " + _intCoilNumber + " ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }
		
        public bool UpdateRCSDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlName, string _strCoilName, int _intCoilNumber, decimal _dcmDC_Deviation
			, decimal _dcmAC_Deviation, decimal _dcmL_Deviation, decimal _dcmC_Deviation, decimal _dcmQ_Deviation, string _strResult)
		{
			bool result = false;

			//----- Sql Open
			conn = new SqlCeConnection(connectionString);
			conn.Open();

			try
			{
				// 커맨드에 커넥션을 연결
				cmd.Connection = conn;

				// 트랜잭션 생성
				tran = conn.BeginTransaction();
				cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbRCSDiagnosisDetail SET ";
                cmd.CommandText = cmd.CommandText + "DC_Deviation = " + _dcmDC_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", AC_Deviation = " + _dcmAC_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", L_Deviation = " + _dcmL_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", C_Deviation = " + _dcmC_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", Q_Deviation = " + _dcmQ_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", Result = '" + _strResult + "' ";
				cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
				cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
				cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
				cmd.CommandText = cmd.CommandText + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
				cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";
				cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strCoilName.Trim() + "' ";
				cmd.CommandText = cmd.CommandText + "AND CoilNumber = " + _intCoilNumber + " ";

				cmd.ExecuteNonQuery();

				// 커밋
				tran.Commit();
				result = true;
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
				tran.Rollback();
				result = false;
			}

			conn.Close();
			return result;
		}

        #endregion

        #endregion


        #region ▣ RCS History

        #region ▣ RCS History - Header

        public int GetRCSDiagnosisHeaderMinOhDegreeInfo(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            int iOh_Degree = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Min(Oh_Degree) as Oh_Degree ";
                _select = _select + "FROM tbTDR_RCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iOh_Degree = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(Regex.Replace(dt.Rows[0]["Oh_Degree"].ToString().Trim(), @"\D", ""));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iOh_Degree;
        }

        public DataTable GetRCSDiagnosisHeaderOhDegreeInfo(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT PlantName, Hogi, Oh_Degree ";
                _select = _select + "FROM tbRCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetRCSDiagnosisHeaderOhDegreeMeasurementDateInfo(string _strPlantName, string _strHogi, string _strOhDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT PlantName, Hogi, Oh_Degree, MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, MeasurementDate ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        #endregion

        #region ▣ RCS History - Detail

        /// <summary>
        /// 주기별/기간병 데이터 가져오기
        /// </summary>
        /// <param name="_strPlantName"></param>
        /// <param name="_strHogi"></param>
        /// <param name="_strOh_Degree"></param>
        /// <param name="_strMeasurementType"></param>
        /// <param name="_strFrequency"></param>
        /// <param name="_strPowerCabinet"></param>
        /// <param name="_strControlName"></param>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        public DataTable GetRCSDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strMeasurementType, string _strFrequency, string _strPowerCabinet, string _strControlName, string _strCoilName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
			dt.Columns.Add("CoilNumber");
			dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
				_select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber";
                _select = _select + ", B.DC_ResistanceValue, B.DC_Deviation, B.AC_ResistanceValue, B.AC_Deviation, B.L_InductanceValue";
                _select = _select + ", B.L_Deviation, B.C_CapacitanceValue, B.C_Deviation, B.Q_FactorValue, B.Q_Deviation, B.Result, B.MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader AS A ";
                _select = _select + "Left Outer Join tbRCSDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName ";
                _select = _select + "AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree ";
                _select = _select + "AND B.PowerCabinet = A.PowerCabinet ";
                _select = _select + "AND B.ControlName = A.ControlName ";
                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOh_Degree.Trim() + "' ";

                if (_strPowerCabinet.Trim() != "전체" && _strPowerCabinet.Trim() != "")
                    _select = _select + "AND A.PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";

                if (_strFrequency.Trim() != "전체" && _strFrequency.Trim() != "")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND B.ControlName = '" + _strControlName.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

				_select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlSeqNumber, B.SeqNumber, B.CoilNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        /// <summary>
        /// 주기별/기간병 측정치 데이터 가져오기
        /// </summary>
        /// <param name="_strPlantName"></param>
        /// <param name="_strHogi"></param>
        /// <param name="_strOh_Degree"></param>
        /// <param name="_strMeasurementType"></param>
        /// <param name="_strFrequency"></param>
        /// <param name="_strPowerCabinet"></param>
        /// <param name="_strControlName"></param>
        /// <param name="_strCoilName"></param>
        /// <param name="_strReferenceHogi"></param>
        /// <param name="_strReferenceOHDegree"></param>
        /// <returns></returns>
        public DataTable GetRCSDiagnosisDetailDataGridViewDataMeasure(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strMeasurementType, string _strFrequency, string _strPowerCabinet, string _strControlName, string _strCoilName
            , string _strReferenceHogi, string _strReferenceOHDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("CoilNumber");
            dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber";
                _select = _select + ", B.DC_ResistanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.DC_ResistanceValue - C1.RdcStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.DC_ResistanceValue - C1.RdcUp_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.DC_ResistanceValue - C1.RdcMove_RCSReferenceValue ELSE 0 END) AS DC_Deviation";
                _select = _select + ", B.AC_ResistanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ELSE 0 END) AS AC_Deviation";
                _select = _select + ", B.L_InductanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ELSE 0 END) AS L_Deviation";
                _select = _select + ", B.C_CapacitanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ELSE 0 END) AS C_Deviation ";
                _select = _select + ", B.Q_FactorValue, (CASE WHEN B.SeqNumber = 1 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ELSE 0 END) AS Q_Deviation ";
                _select = _select + ", B.Result, B.MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader AS A ";
                _select = _select + "LEFT OUTER JOIN tbRCSDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName ";
                _select = _select + "AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree ";
                _select = _select + "AND B.PowerCabinet = A.PowerCabinet ";
                _select = _select + "AND B.ControlName = A.ControlName ";

                _select = _select + "LEFT OUTER JOIN tbRCSReferenceValue AS C1 ";
                _select = _select + "ON C1.PlantName = A.PlantName ";
                _select = _select + "AND C1.Hogi = '" + _strReferenceHogi.Trim() + "' ";
                _select = _select + "AND C1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' ";

                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOh_Degree.Trim() + "' ";

                if (_strPowerCabinet.Trim() != "전체" && _strPowerCabinet.Trim() != "")
                    _select = _select + "AND A.PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";

                if (_strFrequency.Trim() != "전체" && _strFrequency.Trim() != "")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND B.ControlName = '" + _strControlName.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlSeqNumber, B.SeqNumber, B.CoilNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        /// <summary>
        /// 주기별/기간병 평균치 데이터 가져오기
        /// </summary>
        /// <param name="_strPlantName"></param>
        /// <param name="_strHogi"></param>
        /// <param name="_strOh_Degree"></param>
        /// <param name="_strMeasurementType"></param>
        /// <param name="_strFrequency"></param>
        /// <param name="_strPowerCabinet"></param>
        /// <param name="_strControlName"></param>
        /// <param name="_strCoilName"></param>
        /// <param name="_strReferenceHogi"></param>
        /// <param name="_strReferenceOHDegree"></param>
        /// <returns></returns>
        public DataTable GetRCSDiagnosisDetailDataGridViewDataAverage(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strMeasurementType, string _strFrequency, string _strPowerCabinet, string _strControlName, string _strCoilName
            , string _strReferenceHogi, string _strReferenceOHDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("CoilNumber");
            dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, AVG(1) AS CoilNumber, B.ControlSeqNumber, B.SeqNumber";
                _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.DC_ResistanceValue - C1.RdcStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.DC_ResistanceValue - C1.RdcUp_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.DC_ResistanceValue - C1.RdcMove_RCSReferenceValue ELSE 0 END)) AS DC_Deviation";
                _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ELSE 0 END)) AS AC_Deviation";
                _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ELSE 0 END)) AS L_Deviation";
                _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ELSE 0 END)) AS C_Deviation ";
                _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ELSE 0 END)) AS Q_Deviation ";
                _select = _select + ", B.Result, B.MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader AS A ";
                _select = _select + "LEFT OUTER JOIN tbRCSDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName ";
                _select = _select + "AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree ";
                _select = _select + "AND B.PowerCabinet = A.PowerCabinet ";
                _select = _select + "AND B.ControlName = A.ControlName ";

                _select = _select + "LEFT OUTER JOIN tbRCSReferenceValue AS C1 ";
                _select = _select + "ON C1.PlantName = A.PlantName ";
                _select = _select + "AND C1.Hogi = '" + _strReferenceHogi.Trim() + "' ";
                _select = _select + "AND C1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' ";

                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOh_Degree.Trim() + "' ";

                if (_strPowerCabinet.Trim() != "전체" && _strPowerCabinet.Trim() != "")
                    _select = _select + "AND A.PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";

                if (_strFrequency.Trim() != "전체" && _strFrequency.Trim() != "")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND B.ControlName = '" + _strControlName.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                _select = _select + "GROUP BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.ControlSeqNumber, B.SeqNumber, B.Result, B.MeasurementDate ";
                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.ControlSeqNumber, B.SeqNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        #endregion

        #endregion

        
        #region ▣ RCS/DRPI Report

		public DataTable GetRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
			, string _strPowerCabinet)
		{
			//----- Sql Open
			conn = new SqlCeConnection(connectionString);
			conn.Open();

			DataTable dt = new DataTable();

			dt.Columns.Add("PlantName");
			dt.Columns.Add("Hogi");
			dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
			dt.Columns.Add("Temperature_ReferenceValue");
			dt.Columns.Add("TemperatureUpDown_ReferenceValue");
			dt.Columns.Add("Frequency");
			dt.Columns.Add("VoltageLevel");
			dt.Columns.Add("MeasurementCount");
			dt.Columns.Add("Item_Rdc");
			dt.Columns.Add("Item_Rac");
			dt.Columns.Add("Item_L");
			dt.Columns.Add("Item_C");
			dt.Columns.Add("Item_Q");
			dt.Columns.Add("MeasurementMode");
			dt.Columns.Add("MeasurementDate");
			dt.Columns.Add("MeasurementResult");

			try
			{
				// 커맨드에 커넥션을 연결
				string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree, PowerCabinet, ControlName, Temperature_ReferenceValue, TemperatureUpDown_ReferenceValue";
				_select = _select + ", Frequency, VoltageLevel, MeasurementCount, Item_Rdc, Item_Rac, Item_L, Item_C, Item_Q";
				_select = _select + ", MeasurementMode, MeasurementDate, MeasurementResult ";
				_select = _select + "FROM tbRCSDiagnosisHeader ";
				_select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
				_select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
				_select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";

				if (_strPowerCabinet.Trim() != "")
					_select = _select + "AND PowerCabinet IN (" + _strPowerCabinet.Trim() + ") ";

				_select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, MeasurementDate Desc ";

				cmd = new SqlCeCommand(_select, conn);
				SqlCeDataReader rdr = cmd.ExecuteReader();

				// DB에서 가져온 Data를 DataSet에 바인딩
				dt.Load(rdr);
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.Print(ex.Message);
			}

			conn.Close();
			return dt;
		}

        public DataTable GetRCSReportDataInfo(string _strPlantName, string _strHogi, string _strOhDegree, string _strFrequency
            , string _strMeasurementType, string _strCoilName, string _strControlName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
			dt.Columns.Add("CoilNumber");
			dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
				_select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber";
                _select = _select + ", B.DC_ResistanceValue, B.DC_Deviation, B.AC_ResistanceValue, B.AC_Deviation, B.L_InductanceValue";
                _select = _select + ", B.L_Deviation, B.C_CapacitanceValue, B.C_Deviation, B.Q_FactorValue, B.Q_Deviation, B.Result, B.MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader AS A ";
                _select = _select + "Left Outer Join tbRCSDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName ";
                _select = _select + "AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree ";
                _select = _select + "AND B.PowerCabinet = A.PowerCabinet ";
                _select = _select + "AND B.ControlName = A.ControlName ";
                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strFrequency.Trim() != "전체" && _strFrequency.Trim() != "")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

				_select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlSeqNumber, B.SeqNumber, B.CoilNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetRCSReportDataMeasure(string _strPlantName, string _strHogi, string _strOhDegree, string _strFrequency
            , string _strMeasurementType, string _strCoilName, string _strControlName, string _strReferenceHogi, string _strReferenceOHDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("CoilNumber");
            dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber";
                _select = _select + ", B.DC_ResistanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.DC_ResistanceValue - C1.RdcStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.DC_ResistanceValue - C1.RdcUp_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.DC_ResistanceValue - C1.RdcMove_RCSReferenceValue ELSE 0 END) AS DC_Deviation";
                _select = _select + ", B.AC_ResistanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ELSE 0 END) AS AC_Deviation";
                _select = _select + ", B.L_InductanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ELSE 0 END) AS L_Deviation";
                _select = _select + ", B.C_CapacitanceValue, (CASE WHEN B.SeqNumber = 1 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ELSE 0 END) AS C_Deviation ";
                _select = _select + ", B.Q_FactorValue, (CASE WHEN B.SeqNumber = 1 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ELSE 0 END) AS Q_Deviation ";
                _select = _select + " , B.Result, B.MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader AS A ";
                _select = _select + "Left Outer Join tbRCSDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName ";
                _select = _select + "AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree ";
                _select = _select + "AND B.PowerCabinet = A.PowerCabinet ";
                _select = _select + "AND B.ControlName = A.ControlName ";

                _select = _select + "LEFT OUTER JOIN tbRCSReferenceValue AS C1 ";
                _select = _select + "ON C1.PlantName = A.PlantName ";
                _select = _select + "AND C1.Hogi = '" + _strReferenceHogi.Trim() + "' ";
                _select = _select + "AND C1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' ";

                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strFrequency.Trim() != "전체" && _strFrequency.Trim() != "")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlSeqNumber, B.SeqNumber, B.CoilNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetRCSReportDataAverage(string _strPlantName, string _strHogi, string _strOhDegree, string _strFrequency
            , string _strMeasurementType, string _strCoilName, string _strControlName, string _strReferenceHogi, string _strReferenceOHDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("CoilNumber");
            dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, AVG(1) AS CoilNumber, B.ControlSeqNumber, B.SeqNumber";
                _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.DC_ResistanceValue - C1.RdcStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.DC_ResistanceValue - C1.RdcUp_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.DC_ResistanceValue - C1.RdcMove_RCSReferenceValue ELSE 0 END)) AS DC_Deviation";
                _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.AC_ResistanceValue - C1.RacStop_RCSReferenceValue ELSE 0 END)) AS AC_Deviation";
                _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.L_InductanceValue - C1.LStop_RCSReferenceValue ELSE 0 END)) AS L_Deviation";
                _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.C_CapacitanceValue - C1.CStop_RCSReferenceValue ELSE 0 END)) AS C_Deviation ";
                _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue, AVG((CASE WHEN B.SeqNumber = 1 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 2 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ";
                _select = _select + " WHEN B.SeqNumber = 3 THEN B.Q_FactorValue - C1.QStop_RCSReferenceValue ELSE 0 END)) AS Q_Deviation ";
                _select = _select + ", B.Result, B.MeasurementDate ";
                _select = _select + "FROM tbRCSDiagnosisHeader AS A ";
                _select = _select + "Left Outer Join tbRCSDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName ";
                _select = _select + "AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree ";
                _select = _select + "AND B.PowerCabinet = A.PowerCabinet ";
                _select = _select + "AND B.ControlName = A.ControlName ";

                _select = _select + "LEFT OUTER JOIN tbRCSReferenceValue AS C1 ";
                _select = _select + "ON C1.PlantName = A.PlantName ";
                _select = _select + "AND C1.Hogi = '" + _strReferenceHogi.Trim() + "' ";
                _select = _select + "AND C1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' ";

                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strFrequency.Trim() != "전체" && _strFrequency.Trim() != "")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                _select = _select + "GROUP BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlName, B.CoilName, B.ControlSeqNumber, B.SeqNumber, B.Result, B.MeasurementDate ";
                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.PowerCabinet, B.ControlSeqNumber, B.SeqNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIReportDataInfo(string _strPlantName, string _strHogi, string _strOhDegree, string _strDRPIGroup, string _strDRPIType
            , string _strFrequency, string _strCoilName, string _strControlName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
			dt.Columns.Add("CoilNumber");
			dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");
            dt.Columns.Add("Temperature_ReferenceValue");
            dt.Columns.Add("Frequency");
            dt.Columns.Add("VoltageLevel");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
				_select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.DRPIGroup, B.DRPIType, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber, B.MeasurementMode";
                _select = _select + ", B.DC_ResistanceValue, B.DC_Deviation, B.AC_ResistanceValue, B.AC_Deviation, B.L_InductanceValue";
                _select = _select + ", B.L_Deviation, B.C_CapacitanceValue, B.C_Deviation, B.Q_FactorValue, B.Q_Deviation, B.Result, B.MeasurementDate ";
                _select = _select + ", A.Temperature_ReferenceValue, A.Frequency, A.VoltageLevel ";
                _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";
                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                    _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                    _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";
                
                if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                    _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                if (_strCoilName.Trim() != "" && _strCoilName.Trim() != "전체")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.ControlSeqNumber, B.DRPIGroup, B.SeqNumber, B.CoilNumber, B.MeasurementMode ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }
            
        #endregion


        #region ▣ DRPI 측정

        #region ▣ DRPI 측정 - Header

        public string GetDRPIDiagnosisHogiMaxOhDegree(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            string strOh_Degree = "";

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Max(Oh_Degree) AS Oh_Degree ";
                _select = _select + "FROM tbDRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                strOh_Degree = dt == null || dt.Rows.Count < 1 ? "제 0 차" : dt.Rows[0]["Oh_Degree"].ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return strOh_Degree;
        }

        public int GetDRPIDiagnosisHeaderDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree, string _strDRPIGroup
            , string _strDRPIType, string _strControlName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbDRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                _select = _select + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";  
                
                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, decimal _dcmTemperature_ReferenceValue
            , decimal _dcmTemperatureUpDown_ReferenceValue, string _strFrequency, string _strVoltageLevel, int _intMeasurementCount
            , int _intItem_Rdc, int _intItem_Rac, int _intItem_L, int _intItem_C, int _intItem_Q, string _strMeasurementMode
            , string _strMeasurementDate, string _strMeasurementResult)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbDRPIDiagnosisHeader(PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, MeasurementMode";
                cmd.CommandText = cmd.CommandText + ", Temperature_ReferenceValue, TemperatureUpDown_ReferenceValue, Frequency, VoltageLevel";
                cmd.CommandText = cmd.CommandText + ", MeasurementCount, Item_Rdc, Item_Rac, Item_L, Item_C, Item_Q";
                cmd.CommandText = cmd.CommandText + ", MeasurementDate, MeasurementResult) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIGroup.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIType.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementMode.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _dcmTemperature_ReferenceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmTemperatureUpDown_ReferenceValue + ", '";
                cmd.CommandText = cmd.CommandText + _strFrequency.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strVoltageLevel.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intMeasurementCount + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_L + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_C + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Q + ", '";
                cmd.CommandText = cmd.CommandText + _strMeasurementDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementResult.Trim() + "') ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool InsertDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, decimal _dcmTemperature_ReferenceValue
            , decimal _dcmTemperatureUpDown_ReferenceValue, string _strFrequency, string _strVoltageLevel, int _intMeasurementCount
            , int _intItem_Rdc, int _intItem_Rac, int _intItem_L, int _intItem_C, int _intItem_Q, string _strMeasurementMode
            , string _strMeasurementDate, string _strMeasurementResult, int _intTake_ID)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbDRPIDiagnosisHeader(PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, MeasurementMode";
                cmd.CommandText = cmd.CommandText + ", Temperature_ReferenceValue, TemperatureUpDown_ReferenceValue, Frequency, VoltageLevel";
                cmd.CommandText = cmd.CommandText + ", MeasurementCount, Item_Rdc, Item_Rac, Item_L, Item_C, Item_Q";
                cmd.CommandText = cmd.CommandText + ", MeasurementDate, MeasurementResult, TAKE_ID) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIGroup.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIType.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementMode.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _dcmTemperature_ReferenceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmTemperatureUpDown_ReferenceValue + ", '";
                cmd.CommandText = cmd.CommandText + _strFrequency.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strVoltageLevel.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intMeasurementCount + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Rac + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_L + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_C + ", ";
                cmd.CommandText = cmd.CommandText + _intItem_Q + ", '";
                cmd.CommandText = cmd.CommandText + _strMeasurementDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementResult.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intTake_ID + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbDRPIDiagnosisHeader ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #region ▣ DRPI 측정 - Detail

        public int GetDRPIDiagnosisDetailDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _intCoilName, int _intCoilNumber)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbDRPIDiagnosisDetail ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                _select = _select + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";
                _select = _select + "AND CoilName = '" + _intCoilName.Trim() + "' ";
                //_select = _select + "AND CoilNumber = " + _intCoilNumber + " ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }
// 20240401 한인석
// 제어봉 별이 아니라 한 번만 호출
        public StringBuilder GetDRPIDiagnosisDetailDataGridViewExist(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType)
        {
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            //int iCount = 0;
            StringBuilder returnVal = new StringBuilder();
            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT distinct ControlName ";
                _select = _select + "FROM tbDRPIDiagnosisDetail ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                _select = _select + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                if (dt != null || dt.Rows.Count > 0)
                {
                    for(int i = 0; i < dt.Rows.Count; i++)
                    {
                        returnVal.Append(dt.Rows[i]["ControlName"].ToString());
                        if(i < dt.Rows.Count - 1)
                        {
                            returnVal.Append(',');
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return returnVal;
        }

        public int GetDRPIDiagnosisDetailDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbDRPIDiagnosisDetail ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                _select = _select + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertDRPIDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strCoilName, int _intCoilNumber, string _strMeasurementMode
            , decimal _dcmDC_ResistanceValue, decimal _dcmDC_Deviation, decimal _dcmAC_ResistanceValue, decimal _dcmAC_Deviation
            , decimal _dcmL_InductanceValue, decimal _dcmL_Deviation, decimal _dcmC_CapacitanceValue, decimal _dcmC_Deviation
            , decimal _dcmQ_FactorValue, decimal _dcmQ_Deviation, string _strResult, string _strMeasurementDate, int _intSeqNumber
			, int _intControlSeqNumber)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbDRPIDiagnosisDetail(PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, CoilName, CoilNumber";
                cmd.CommandText = cmd.CommandText + ", MeasurementMode, DC_ResistanceValue, DC_Deviation, AC_ResistanceValue, AC_Deviation";
                cmd.CommandText = cmd.CommandText + ", L_InductanceValue, L_Deviation, C_CapacitanceValue, C_Deviation, Q_FactorValue, Q_Deviation";
				cmd.CommandText = cmd.CommandText + ", Result, MeasurementDate, SeqNumber, ControlSeqNumber) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIGroup.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIType.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strCoilName.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intCoilNumber + ", '";
                cmd.CommandText = cmd.CommandText + _strMeasurementMode.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _dcmDC_ResistanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmDC_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmAC_ResistanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmAC_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmL_InductanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmL_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmC_CapacitanceValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmC_Deviation + ", ";
                cmd.CommandText = cmd.CommandText + _dcmQ_FactorValue + ", ";
                cmd.CommandText = cmd.CommandText + _dcmQ_Deviation + ", '";
                cmd.CommandText = cmd.CommandText + _strResult + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementDate.Trim() + "', ";
				cmd.CommandText = cmd.CommandText + _intSeqNumber + ", ";
				cmd.CommandText = cmd.CommandText + _intControlSeqNumber + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteDRPIDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbDRPIDiagnosisDetail ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteDRPIDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strCoilName, int _intCoilNumber)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbDRPIDiagnosisDetail ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strCoilName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilNumber = " + _intCoilNumber + " ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #endregion


        #region ▣ DRPI History

        #region ▣ DRPI History - Header

        public int GetDRPIDiagnosisHeaderMinOhDegreeInfo(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            int iOh_Degree = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Min(Oh_Degree) as Oh_Degree ";
                _select = _select + "FROM tbTDR_DRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iOh_Degree = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(Regex.Replace(dt.Rows[0]["Oh_Degree"].ToString().Trim(), @"\D", ""));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iOh_Degree;
        }

        public DataTable GetDRPIDiagnosisHeaderOhDegreeInfo(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT PlantName, Hogi, Oh_Degree ";
                _select = _select + "FROM tbDRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIDiagnosisHeaderOhDegreeMeasurementDateInfo(string _strPlantName, string _strHogi, string _strOhDegree, string _strDRPIGroup)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("MeasurementDate");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT DISTINCT PlantName, Hogi, Oh_Degree, DRPIGroup, MeasurementDate ";
                _select = _select + "FROM tbDRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "전체" && _strDRPIGroup.Trim() != "")
                    _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, DRPIGroup, MeasurementDate ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOhDegreeFrom, string _strDRPIGroup
            , string _strMeasurementDate, string _strMeasurementType, string _strType, string _strControlRod, string _strCoilName, string _strFrequency)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("Temperature_ReferenceValue");
            dt.Columns.Add("TemperatureUpDown_ReferenceValue");
            dt.Columns.Add("Frequency");
            dt.Columns.Add("VoltageLevel");
            dt.Columns.Add("StandardResistance");
            dt.Columns.Add("MeasurementCount");
            dt.Columns.Add("Item_Rdc");
            dt.Columns.Add("Item_Rac");
            dt.Columns.Add("Item_L");
            dt.Columns.Add("Item_C");
            dt.Columns.Add("Item_Q");
            dt.Columns.Add("MeasurementDate");
            dt.Columns.Add("MeasurementResult");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, MeasurementMode, Temperature_ReferenceValue, TemperatureUpDown_ReferenceValue";
                _select = _select + ", Frequency, VoltageLevel, StandardResistance, MeasurementCount, Item_Rdc, Item_Rac, Item_L, Item_C, Item_Q";
                _select = _select + ", MeasurementDate, MeasurementResult ";
                _select = _select + "FROM tbDRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOhDegreeFrom.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "전체" && _strDRPIGroup.Trim() != "")
                    _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                if (_strType.Trim() != "전체" && _strType.Trim() != "")
                    _select = _select + "AND DRPIType = '" + _strType.Trim() + "' ";

                if (_strControlRod.Trim() != "전체" && _strControlRod.Trim() != "")
                    _select = _select + "AND ControlName = '" + _strControlRod.Trim() + "' ";

                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, DRPIGroup ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        #endregion

        #region ▣ DRPI History - Detail

        public DataTable GetDRPIDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOhDegree, string _strDRPIGroup
            , string _strDRPIType, string _strControlName, string _strCoilName, string _strFrequency)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
			dt.Columns.Add("CoilNumber");
			dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");
            dt.Columns.Add("Temperature_ReferenceValue");
            dt.Columns.Add("Frequency");
            dt.Columns.Add("VoltageLevel");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
				_select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.DRPIGroup, B.DRPIType, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber, B.MeasurementMode";
                _select = _select + ", B.DC_ResistanceValue, B.DC_Deviation, B.AC_ResistanceValue, B.AC_Deviation, B.L_InductanceValue";
                _select = _select + ", B.L_Deviation, B.C_CapacitanceValue, B.C_Deviation, B.Q_FactorValue, B.Q_Deviation, B.Result, B.MeasurementDate ";
                _select = _select + ", A.Temperature_ReferenceValue, A.Frequency, A.VoltageLevel ";
                _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";
                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                    _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                    _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                    _select = _select + "AND B.ControlName = '" + _strControlName.Trim() + "' ";

                if (_strCoilName.Trim() != "" && _strCoilName.Trim() != "전체")
                    _select = _select + "AND B.CoilName = '" + _strCoilName.Trim() + "' ";

                if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.ControlSeqNumber, B.DRPIGroup, B.SeqNumber, B.CoilNumber, B.MeasurementMode ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIDiagnosisDetailDataGridViewDataMeasure(string _strPlantName, string _strHogi, string _strOhDegree, string _strDRPIGroup
            , string _strDRPIType, string _strControlName, string _strCoilName, string _strFrequency, string _strReferenceHogi, string _strReferenceOHDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("CoilNumber");
            dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");
            dt.Columns.Add("Temperature_ReferenceValue");
            dt.Columns.Add("Frequency");
            dt.Columns.Add("VoltageLevel");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.DRPIGroup, B.DRPIType, B.ControlName, B.CoilName, B.CoilNumber, B.ControlSeqNumber, B.SeqNumber, B.MeasurementMode ";
                _select = _select + ", B.DC_ResistanceValue ";
                _select = _select + ", (CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.DC_ResistanceValue - A1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.DC_ResistanceValue - A2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - A3.Rdc_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.DC_ResistanceValue - B1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.DC_ResistanceValue - B2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - B3.Rdc_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.DC_ResistanceValue - A1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.DC_ResistanceValue - A2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - A3.Rdc_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.DC_ResistanceValue - B1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.DC_ResistanceValue - B2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - B3.Rdc_DRPIReferenceValue END) END) AS DC_Deviation ";
                _select = _select + ", B.AC_ResistanceValue ";
                _select = _select + ", (CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.AC_ResistanceValue - A1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.AC_ResistanceValue - A2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - A3.Rac_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.AC_ResistanceValue - B1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.AC_ResistanceValue - B2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - B3.Rac_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.AC_ResistanceValue - A1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.AC_ResistanceValue - A2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - A3.Rac_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.AC_ResistanceValue - B1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.AC_ResistanceValue - B2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - B3.Rac_DRPIReferenceValue END) END) AS AC_Deviation ";
                _select = _select + ", B.L_InductanceValue ";
                _select = _select + ", (CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.L_InductanceValue - A1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.L_InductanceValue - A2.L_DRPIReferenceValue ELSE B.L_InductanceValue - A3.L_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.L_InductanceValue - B1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.L_InductanceValue - B2.L_DRPIReferenceValue ELSE B.L_InductanceValue - B3.L_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.L_InductanceValue - A1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.L_InductanceValue - A2.L_DRPIReferenceValue ELSE B.L_InductanceValue - A3.L_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.L_InductanceValue - B1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.L_InductanceValue - B2.L_DRPIReferenceValue ELSE B.L_InductanceValue - B3.L_DRPIReferenceValue END) END) AS L_Deviation ";
                _select = _select + ", B.C_CapacitanceValue ";
                _select = _select + ", (CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.C_CapacitanceValue - A1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.C_CapacitanceValue - A2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - A3.C_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.C_CapacitanceValue - B1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.C_CapacitanceValue - B2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - B3.C_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.C_CapacitanceValue - A1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.C_CapacitanceValue - A2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - A3.C_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.C_CapacitanceValue - B1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.C_CapacitanceValue - B2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - B3.C_DRPIReferenceValue END) END) AS C_Deviation ";
                _select = _select + ", B.Q_FactorValue ";
                _select = _select + ", (CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.Q_FactorValue - A1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.Q_FactorValue - A2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - A3.Q_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.Q_FactorValue - B1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.Q_FactorValue - B2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - B3.Q_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.Q_FactorValue - A1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.Q_FactorValue - A2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - A3.Q_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.Q_FactorValue - B1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.Q_FactorValue - B2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - B3.Q_DRPIReferenceValue END) END) AS Q_Deviation ";
                _select = _select + ", B.Result, B.MeasurementDate, A.Temperature_ReferenceValue, A.Frequency, A.VoltageLevel ";
                _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";

                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS A1 ON A1.PlantName = A.PlantName AND A1.Hogi = '" + _strReferenceHogi.Trim() + "' AND A1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND A1.CoilCabinetType = B.DRPIType AND A1.CoilCabinetName = 'A1' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS A2 ON A2.PlantName = A.PlantName AND A2.Hogi = '" + _strReferenceHogi.Trim() + "' AND A2.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND A2.CoilCabinetType = B.DRPIType AND A2.CoilCabinetName = 'A2' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS A3 ON A3.PlantName = A.PlantName AND A3.Hogi = '" + _strReferenceHogi.Trim() + "' AND A3.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND A3.CoilCabinetType = B.DRPIType AND A3.CoilCabinetName LIKE 'A3%' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS B1 ON B1.PlantName = A.PlantName AND B1.Hogi = '" + _strReferenceHogi.Trim() + "' AND B1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND B1.CoilCabinetType = B.DRPIType AND B1.CoilCabinetName = 'B1' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS B2 ON B2.PlantName = A.PlantName AND B2.Hogi = '" + _strReferenceHogi.Trim() + "' AND B2.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND B2.CoilCabinetType = B.DRPIType AND B2.CoilCabinetName = 'B2' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS B3 ON B3.PlantName = A.PlantName AND B3.Hogi = '" + _strReferenceHogi.Trim() + "' AND B3.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND B3.CoilCabinetType = B.DRPIType AND B3.CoilCabinetName LIKE 'B3%' ";

                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                    _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                    _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                    _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                if (_strCoilName.Trim() != "" && _strCoilName.Trim() != "전체")
                    _select = _select + "AND B.CoilName IN (" + _strCoilName.Trim() + ") ";

                if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.ControlSeqNumber, B.DRPIGroup, B.SeqNumber, B.MeasurementMode ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIDiagnosisDetailDataGridViewDataAverage(string _strPlantName, string _strHogi, string _strOhDegree, string _strDRPIGroup
            , string _strDRPIType, string _strControlName, string _strCoilName, string _strFrequency, string _strReferenceHogi, string _strReferenceOHDegree)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("CoilNumber");
            dt.Columns.Add("ControlSeqNumber");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("MeasurementMode");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("DC_Deviation");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("AC_Deviation");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("L_Deviation");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("C_Deviation");
            dt.Columns.Add("Q_FactorValue");
            dt.Columns.Add("Q_Deviation");
            dt.Columns.Add("Result");
            dt.Columns.Add("MeasurementDate");
            dt.Columns.Add("Temperature_ReferenceValue");
            dt.Columns.Add("Frequency");
            dt.Columns.Add("VoltageLevel");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT B.PlantName, B.Hogi, B.Oh_Degree, B.DRPIGroup, B.DRPIType, B.ControlName, B.CoilName, AVG(1) AS CoilNumber, B.ControlSeqNumber, B.SeqNumber, B.MeasurementMode ";
                _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue ";
                _select = _select + ", AVG((CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.DC_ResistanceValue - A1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.DC_ResistanceValue - A2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - A3.Rdc_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.DC_ResistanceValue - B1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.DC_ResistanceValue - B2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - B3.Rdc_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.DC_ResistanceValue - A1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.DC_ResistanceValue - A2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - A3.Rdc_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.DC_ResistanceValue - B1.Rdc_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.DC_ResistanceValue - B2.Rdc_DRPIReferenceValue ELSE B.DC_ResistanceValue - B3.Rdc_DRPIReferenceValue END) END)) AS DC_Deviation ";
                _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue ";
                _select = _select + ", AVG((CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.AC_ResistanceValue - A1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.AC_ResistanceValue - A2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - A3.Rac_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.AC_ResistanceValue - B1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.AC_ResistanceValue - B2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - B3.Rac_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.AC_ResistanceValue - A1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.AC_ResistanceValue - A2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - A3.Rac_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.AC_ResistanceValue - B1.Rac_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.AC_ResistanceValue - B2.Rac_DRPIReferenceValue ELSE B.AC_ResistanceValue - B3.Rac_DRPIReferenceValue END) END)) AS AC_Deviation ";
                _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue ";
                _select = _select + ", AVG((CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.L_InductanceValue - A1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.L_InductanceValue - A2.L_DRPIReferenceValue ELSE B.L_InductanceValue - A3.L_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.L_InductanceValue - B1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.L_InductanceValue - B2.L_DRPIReferenceValue ELSE B.L_InductanceValue - B3.L_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.L_InductanceValue - A1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.L_InductanceValue - A2.L_DRPIReferenceValue ELSE B.L_InductanceValue - A3.L_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.L_InductanceValue - B1.L_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.L_InductanceValue - B2.L_DRPIReferenceValue ELSE B.L_InductanceValue - B3.L_DRPIReferenceValue END) END)) AS L_Deviation ";
                _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue ";
                _select = _select + ", AVG((CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.C_CapacitanceValue - A1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.C_CapacitanceValue - A2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - A3.C_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.C_CapacitanceValue - B1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.C_CapacitanceValue - B2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - B3.C_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.C_CapacitanceValue - A1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.C_CapacitanceValue - A2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - A3.C_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.C_CapacitanceValue - B1.C_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.C_CapacitanceValue - B2.C_DRPIReferenceValue ELSE B.C_CapacitanceValue - B3.C_DRPIReferenceValue END) END)) AS C_Deviation ";
                _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue ";
                _select = _select + ", AVG((CASE WHEN B.DRPIGroup = 'A' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.Q_FactorValue - A1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.Q_FactorValue - A2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - A3.Q_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '제어용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.Q_FactorValue - B1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.Q_FactorValue - B2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - B3.Q_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'A' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'A1' THEN B.Q_FactorValue - A1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'A2' THEN B.Q_FactorValue - A2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - A3.Q_DRPIReferenceValue END) ";
                _select = _select + "        WHEN B.DRPIGroup = 'B' AND B.DRPIType = '정지용' THEN (CASE WHEN B.CoilName = 'B1' THEN B.Q_FactorValue - B1.Q_DRPIReferenceValue ";
                _select = _select + "             WHEN B.CoilName = 'B2' THEN B.Q_FactorValue - B2.Q_DRPIReferenceValue ELSE B.Q_FactorValue - B3.Q_DRPIReferenceValue END) END)) AS Q_Deviation ";
                _select = _select + ", B.Result, B.MeasurementDate, A.Temperature_ReferenceValue, A.Frequency, A.VoltageLevel ";
                _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";

                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS A1 ON A1.PlantName = A.PlantName AND A1.Hogi = '" + _strReferenceHogi.Trim() + "' AND A1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND A1.CoilCabinetType = B.DRPIType AND A1.CoilCabinetName = 'A1' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS A2 ON A2.PlantName = A.PlantName AND A2.Hogi = '" + _strReferenceHogi.Trim() + "' AND A2.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND A2.CoilCabinetType = B.DRPIType AND A2.CoilCabinetName = 'A2' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS A3 ON A3.PlantName = A.PlantName AND A3.Hogi = '" + _strReferenceHogi.Trim() + "' AND A3.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND A3.CoilCabinetType = B.DRPIType AND A3.CoilCabinetName LIKE 'A3%' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS B1 ON B1.PlantName = A.PlantName AND B1.Hogi = '" + _strReferenceHogi.Trim() + "' AND B1.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND B1.CoilCabinetType = B.DRPIType AND B1.CoilCabinetName = 'B1' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS B2 ON B2.PlantName = A.PlantName AND B2.Hogi = '" + _strReferenceHogi.Trim() + "' AND B2.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND B2.CoilCabinetType = B.DRPIType AND B2.CoilCabinetName = 'B2' ";
                _select = _select + "LEFT OUTER JOIN tbDRPIReferenceValue AS B3 ON B3.PlantName = A.PlantName AND B3.Hogi = '" + _strReferenceHogi.Trim() + "' AND B3.Oh_Degree = '" + _strReferenceOHDegree.Trim() + "' AND B3.CoilCabinetType = B.DRPIType AND B3.CoilCabinetName LIKE 'B3%' ";

                _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                    _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                    _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                    _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                if (_strCoilName.Trim() != "" && _strCoilName.Trim() != "전체")
                    _select = _select + "AND B.CoilName IN (" + _strCoilName.Trim() + ") ";

                if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                    _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                _select = _select + "GROUP BY B.PlantName, B.Hogi, B.Oh_Degree, B.DRPIGroup, B.DRPIType, B.ControlName, B.CoilName, B.ControlSeqNumber, B.SeqNumber, B.MeasurementMode ";
                _select = _select + ", B.Result, B.MeasurementDate, A.Temperature_ReferenceValue, A.Frequency, A.VoltageLevel ";
                _select = _select + "ORDER BY B.PlantName, B.Hogi, B.Oh_Degree, B.ControlSeqNumber, B.DRPIGroup, B.SeqNumber, B.MeasurementMode ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public DataTable GetDRPIDiagnosisDetailDataGridViewDataAverage(string _strPlantName, string _strHogi, string _strOhDegree, string _strDRPIGroup
            , string _strDRPIType, string _strControlName, string _strCoilName, string _strFrequency)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("SeqNumber");
            dt.Columns.Add("DC_ResistanceValue");
            dt.Columns.Add("AC_ResistanceValue");
            dt.Columns.Add("L_InductanceValue");
            dt.Columns.Add("C_CapacitanceValue");
            dt.Columns.Add("Q_FactorValue");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                if (_strCoilName.Trim() != "" && _strCoilName.Trim() != "전체")
                {
                    _select = "SELECT B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                    _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue ";
                    _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue ";
                    _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue ";
                    _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue ";
                    _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue ";
                    _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                    _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                    _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                    _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                    _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";

                    _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                    _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                    _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";

                    if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                        _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                    if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                        _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                    if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                        _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                    if (_strCoilName.Trim() != "" && _strCoilName.Trim() != "전체")
                        _select = _select + "AND B.CoilName IN (" + _strCoilName.Trim() + ") ";

                    if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                        _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                    _select = _select + "GROUP BY B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                    _select = _select + "ORDER BY B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                }
                else
                {
                    _select = "SELECT B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                    _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue ";
                    _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue ";
                    _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue ";
                    _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue ";
                    _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue ";
                    _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                    _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                    _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                    _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                    _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";

                    _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                    _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                    _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                    _select = _select + "AND B.SeqNumber = 1 ";

                    if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                        _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                    if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                        _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                    if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                        _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                    if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                        _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                    _select = _select + "GROUP BY B.DRPIGroup, B.DRPIType, B.SeqNumber ";

                    _select = _select + " UNION ALL ";

                    _select = _select + "SELECT B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                    _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue ";
                    _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue ";
                    _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue ";
                    _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue ";
                    _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue ";
                    _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                    _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                    _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                    _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                    _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";

                    _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                    _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                    _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                    _select = _select + "AND B.SeqNumber = 2 ";

                    if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                        _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                    if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                        _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                    if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                        _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                    if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                        _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                    _select = _select + "GROUP BY B.DRPIGroup, B.DRPIType, B.SeqNumber ";

                    _select = _select + " UNION ALL ";

                    _select = _select + "SELECT B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                    _select = _select + ", AVG(B.DC_ResistanceValue) AS DC_ResistanceValue ";
                    _select = _select + ", AVG(B.AC_ResistanceValue) AS AC_ResistanceValue ";
                    _select = _select + ", AVG(B.L_InductanceValue) AS L_InductanceValue ";
                    _select = _select + ", AVG(B.C_CapacitanceValue) AS C_CapacitanceValue ";
                    _select = _select + ", AVG(B.Q_FactorValue) AS Q_FactorValue ";

                    _select = _select + "FROM (";
                    _select = _select + "SELECT B.DRPIGroup, B.DRPIType, 3 AS SeqNumber ";
                    _select = _select + ", B.DC_ResistanceValue, B.AC_ResistanceValue ";
                    _select = _select + ", B.L_InductanceValue, B.C_CapacitanceValue ";
                    _select = _select + ", B.Q_FactorValue ";
                    _select = _select + "FROM tbDRPIDiagnosisHeader AS A ";
                    _select = _select + "LEFT Outer JOIN tbDRPIDiagnosisDetail AS B ";
                    _select = _select + "ON B.PlantName = A.PlantName AND B.Hogi = A.Hogi ";
                    _select = _select + "AND B.Oh_Degree = A.Oh_Degree AND B.DRPIGroup = A.DRPIGroup ";
                    _select = _select + "AND B.DRPIType = A.DRPIType AND B.ControlName = A.ControlName ";

                    _select = _select + "WHERE A.PlantName = '" + _strPlantName.Trim() + "' ";
                    _select = _select + "AND A.Hogi = '" + _strHogi.Trim() + "' ";
                    _select = _select + "AND A.Oh_Degree = '" + _strOhDegree.Trim() + "' ";
                    _select = _select + "AND B.SeqNumber > 2 ";

                    if (_strDRPIGroup.Trim() != "" && _strDRPIGroup.Trim() != "전체")
                        _select = _select + "AND A.DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                    if (_strDRPIType.Trim() != "" && _strDRPIType.Trim() != "전체")
                        _select = _select + "AND B.DRPIType = '" + _strDRPIType.Trim() + "' ";

                    if (_strControlName.Trim() != "" && _strControlName.Trim() != "전체")
                        _select = _select + "AND B.ControlName IN (" + _strControlName.Trim() + ") ";

                    if (_strFrequency.Trim() != "" && _strFrequency.Trim() != "전체")
                        _select = _select + "AND A.Frequency = '" + _strFrequency.Trim() + "' ";

                    _select = _select + ") AS B ";
                    _select = _select + "GROUP BY B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                    _select = _select + "ORDER BY B.DRPIGroup, B.DRPIType, B.SeqNumber ";
                }

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public bool UpdateDRPIDiagnosisDetailDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strCoilName, int _intCoilNumber, decimal _dcmDC_Deviation
            , decimal _dcmAC_Deviation, decimal _dcmL_Deviation, decimal _dcmC_Deviation, decimal _dcmQ_Deviation, string _strResult)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbDRPIDiagnosisDetail SET ";
                cmd.CommandText = cmd.CommandText + "DC_Deviation = " + _dcmDC_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", AC_Deviation = " + _dcmAC_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", L_Deviation = " + _dcmL_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", C_Deviation = " + _dcmC_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", Q_Deviation = " + _dcmQ_Deviation + " ";
                cmd.CommandText = cmd.CommandText + ", Result = '" + _strResult + "' ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strCoilName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilNumber = " + _intCoilNumber + " ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion

        #endregion


        #region ▣ TDR-RCS 측정/이력

        /// <summary>
        /// 주기별/기간병 데이터 가져오기
        /// </summary>
        /// <param name="_strPlantName"></param>
        /// <param name="_strHogi"></param>
        /// <param name="_strOh_Degree"></param>
        /// <param name="_strMeasurementType"></param>
        /// <param name="_strFrequency"></param>
        /// <param name="_strPowerCabinet"></param>
        /// <param name="_strControlName"></param>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        public DataTable GetTDRRCSDiagnosisDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOhDegreeF, string _strOhDegreeT
            , string _strPowerCabinet, string _strControlName, string _strCoilName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("PowerCabinet");
            dt.Columns.Add("ControlRodName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("Measurement_StartDate");
            dt.Columns.Add("Measurement_EndDate");
            dt.Columns.Add("ImageDataPathFileName");
            dt.Columns.Add("TextDataPathFileName");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree, PowerCabinet, ControlRodName, CoilName";
                _select = _select + ", Measurement_StartDate, Measurement_EndDate, ImageDataPathFileName, TextDataPathFileName ";
                _select = _select + "FROM tbTDR_RCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                if (_strOhDegreeF.Trim() != "" && _strOhDegreeT.Trim() != "")
                    _select = _select + "AND Oh_Degree BetWeen '" + _strOhDegreeF.Trim() + "' AND '" + _strOhDegreeT.Trim() + "' ";
                else if (_strOhDegreeF.Trim() != "" && _strOhDegreeT.Trim() == "")
                    _select = _select + "AND Oh_Degree = '" + _strOhDegreeF.Trim() + "' ";
                else if (_strOhDegreeF.Trim() == "" && _strOhDegreeT.Trim() != "")
                    _select = _select + "AND Oh_Degree = '" + _strOhDegreeT.Trim() + "' ";

                if (_strPowerCabinet.Trim() != "전체" && _strPowerCabinet.Trim() != "")
                    _select = _select + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND ControlRodName = '" + _strControlName.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND CoilName = '" + _strCoilName.Trim() + "' ";

                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, PowerCabinet, ControlRodName, SeqNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public string GetTDRRCSDiagnosisHogiMaxOhDegree(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            string strOh_Degree = "";

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Max(Oh_Degree) AS Oh_Degree ";
                _select = _select + "FROM tbTDR_RCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                strOh_Degree = dt == null || dt.Rows.Count < 1 ? "제 0 차" : dt.Rows[0]["Oh_Degree"].ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return strOh_Degree;
        }

        public int GetTDRRCSDiagnosisHeaderDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strPowerCabinet, string _strControlRodName, string _strColilName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbTDR_RCSDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                _select = _select + "AND ControlRodName = '" + _strControlRodName.Trim() + "' ";
                _select = _select + "AND CoilName = '" + _strColilName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertTDRRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi
            , string _strOh_Degree, string _strPowerCabinet, string _strControlRodName, string _strColilName
            , string _strMeasurementStartDate, string _strMeasurementEndDate, string _strImageDataPathFileName
            , string _strTextDataPathFileName, int _intSeqNumber)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbTDR_RCSDiagnosisHeader(PlantName, Hogi, Oh_Degree, PowerCabinet, ControlRodName, CoilName";
                cmd.CommandText = cmd.CommandText + ", Measurement_StartDate, Measurement_EndDate, ImageDataPathFileName, TextDataPathFileName, SeqNumber) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strPowerCabinet.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlRodName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strColilName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementStartDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementEndDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strImageDataPathFileName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strTextDataPathFileName.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intSeqNumber + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteTDRRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi
            , string _strOh_Degree, string _strPowerCabinet, string _strControlRodName, string _strColilName)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbTDR_RCSDiagnosisHeader ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlRodName = '" + _strControlRodName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strColilName.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool UpDateTDRRCSDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi
            , string _strOh_Degree, string _strPowerCabinet, string _strControlRodName, string _strColilName
            , string _strMeasurementEndDate)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbTDR_RCSDiagnosisHeader SET ";
                cmd.CommandText = cmd.CommandText + "Measurement_EndDate = '" + _strMeasurementEndDate.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND PowerCabinet = '" + _strPowerCabinet.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlRodName = '" + _strControlRodName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strColilName.Trim() + "' ";
                
                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion


        #region ▣ TDR-DRPI 측정/이력

        /// <summary>
        /// 주기별/기간병 데이터 가져오기
        /// </summary>
        /// <param name="_strPlantName"></param>
        /// <param name="_strHogi"></param>
        /// <param name="_strOh_Degree"></param>
        /// <param name="_strMeasurementType"></param>
        /// <param name="_strFrequency"></param>
        /// <param name="_strDRPIGroup"></param>
        /// <param name="_strDRPIType"></param>
        /// <param name="_strControlName"></param>
        /// <param name="_strCoilName"></param>
        /// <returns></returns>
        public DataTable GetTDRDRPIDiagnosisDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOhDegreeF, string _strOhDegreeT
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strCoilName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("PlantName");
            dt.Columns.Add("Hogi");
            dt.Columns.Add("Oh_Degree");
            dt.Columns.Add("DRPIGroup");
            dt.Columns.Add("DRPIType");
            dt.Columns.Add("ControlName");
            dt.Columns.Add("CoilName");
            dt.Columns.Add("Measurement_StartDate");
            dt.Columns.Add("Measurement_EndDate");
            dt.Columns.Add("ImageDataPathFileName");
            dt.Columns.Add("TextDataPathFileName");

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, CoilName";
                _select = _select + ", Measurement_StartDate, Measurement_EndDate, ImageDataPathFileName, TextDataPathFileName ";
                _select = _select + "FROM tbTDR_DRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                if (_strOhDegreeF.Trim() != "" && _strOhDegreeT.Trim() != "")
                    _select = _select + "AND Oh_Degree BetWeen '" + _strOhDegreeF.Trim() + "' AND '" + _strOhDegreeT.Trim() + "' ";
                else if (_strOhDegreeF.Trim() != "" && _strOhDegreeT.Trim() == "")
                    _select = _select + "AND Oh_Degree = '" + _strOhDegreeF.Trim() + "' ";
                else if (_strOhDegreeF.Trim() == "" && _strOhDegreeT.Trim() != "")
                    _select = _select + "AND Oh_Degree = '" + _strOhDegreeT.Trim() + "' ";

                if (_strDRPIGroup.Trim() != "전체" && _strDRPIGroup.Trim() != "")
                    _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";

                if (_strDRPIType.Trim() != "전체" && _strDRPIType.Trim() != "")
                    _select = _select + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";

                if (_strControlName.Trim() != "전체" && _strControlName.Trim() != "")
                    _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";

                if (_strCoilName.Trim() != "전체" && _strCoilName.Trim() != "")
                    _select = _select + "AND CoilName = '" + _strCoilName.Trim() + "' ";

                _select = _select + "ORDER BY PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, SeqNumber ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return dt;
        }

        public string GetTDRDRPIDiagnosisHogiMaxOhDegree(string _strPlantName, string _strHogi)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("Oh_Degree");

            string strOh_Degree = "";

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Max(Oh_Degree) AS Oh_Degree ";
                _select = _select + "FROM tbTDR_DRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                strOh_Degree = dt == null || dt.Rows.Count < 1 ? "제 0 차" : dt.Rows[0]["Oh_Degree"].ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return strOh_Degree;
        }

        public int GetTDRDRPIDiagnosisHeaderDataGridViewDataCount(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strColilName)
        {
            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            DataTable dt = new DataTable();

            dt.Columns.Add("iCount");

            int iCount = 0;

            try
            {
                // 커맨드에 커넥션을 연결
                string _select = "";
                _select = "SELECT Count(*) AS iCount ";
                _select = _select + "FROM tbTDR_DRPIDiagnosisHeader ";
                _select = _select + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                _select = _select + "AND Hogi = '" + _strHogi.Trim() + "' ";
                _select = _select + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                _select = _select + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                _select = _select + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                _select = _select + "AND ControlName = '" + _strControlName.Trim() + "' ";
                _select = _select + "AND CoilName = '" + _strColilName.Trim() + "' ";

                cmd = new SqlCeCommand(_select, conn);
                SqlCeDataReader rdr = cmd.ExecuteReader();

                // DB에서 가져온 Data를 DataSet에 바인딩
                dt.Load(rdr);

                iCount = dt == null || dt.Rows.Count < 1 ? 0 : Convert.ToInt32(dt.Rows[0]["iCount"].ToString().Trim());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }

            conn.Close();
            return iCount;
        }

        public bool InsertTDRDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strColilName, string _strMeasurementStartDate
            , string _strMeasurementEndDate, string _strImageDataPathFileName, string _strTextDataPathFileName, int _intSeqNumber)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "INSERT INTO tbTDR_DRPIDiagnosisHeader(PlantName, Hogi, Oh_Degree, DRPIGroup, DRPIType, ControlName, CoilName";
                cmd.CommandText = cmd.CommandText + ", Measurement_StartDate, Measurement_EndDate, ImageDataPathFileName, TextDataPathFileName, SeqNumber) ";
                cmd.CommandText = cmd.CommandText + "VALUES('";
                cmd.CommandText = cmd.CommandText + _strPlantName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strHogi.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strOh_Degree.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIGroup.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strDRPIType.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strControlName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strColilName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementStartDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strMeasurementEndDate.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strImageDataPathFileName.Trim() + "', '";
                cmd.CommandText = cmd.CommandText + _strTextDataPathFileName.Trim() + "', ";
                cmd.CommandText = cmd.CommandText + _intSeqNumber + ") ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool DeleteTDRDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strColilName)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "DELETE tbTDR_DRPIDiagnosisHeader ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strColilName.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        public bool UpDateTDRDRPIDiagnosisHeaderDataGridViewDataInfo(string _strPlantName, string _strHogi, string _strOh_Degree
            , string _strDRPIGroup, string _strDRPIType, string _strControlName, string _strColilName, string _strMeasurementEndDate)
        {
            bool result = false;

            //----- Sql Open
            conn = new SqlCeConnection(connectionString);
            conn.Open();

            try
            {
                // 커맨드에 커넥션을 연결
                cmd.Connection = conn;

                // 트랜잭션 생성
                tran = conn.BeginTransaction();
                cmd.Transaction = tran;

                cmd.CommandText = "UPDATE tbTDR_DRPIDiagnosisHeader SET ";
                cmd.CommandText = cmd.CommandText + "Measurement_EndDate = '" + _strMeasurementEndDate.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "WHERE PlantName = '" + _strPlantName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Hogi = '" + _strHogi.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND Oh_Degree = '" + _strOh_Degree.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIGroup = '" + _strDRPIGroup.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND DRPIType = '" + _strDRPIType.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND ControlName = '" + _strControlName.Trim() + "' ";
                cmd.CommandText = cmd.CommandText + "AND CoilName = '" + _strColilName.Trim() + "' ";

                cmd.ExecuteNonQuery();

                // 커밋
                tran.Commit();
                result = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                tran.Rollback();
                result = false;
            }

            conn.Close();
            return result;
        }

        #endregion
    }


    public class FunctionMSAccessDataControl
    {
        // 변수 선언
        private OleDbConnection con;
        private OleDbDataAdapter adap;
        private DataTable dataTable;

        /// <summary>
        /// OleDbConnection 객체 생성 및 Open
        /// </summary>
        public void OpenOleDB(string _strAccessDBPath, string _strAccessDBFileName)
        {
            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}{1}", _strAccessDBPath.Trim(), _strAccessDBFileName.Trim());

                con.Open();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// OleDbConnection 객체 Close
        /// </summary>
        public void CloseOleDB()
        {
            try
            {
                // OleDbConnection 객체 Close
                con.Close();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
        }

        /// <summary>
        /// Access DB 데이터 가져오기
        /// </summary>
        /// <param name="_strQuery"></param>
        /// <param name="_intColumnCount"></param>
        /// <returns></returns>
        public DataTable SelectOleDb(string _strAccessDBFileName, string _strQuery)
        {
            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}", _strAccessDBFileName.Trim());
                con.Open();

                // OleDbCommand 객체 생성
                OleDbCommand cmdSelect = new OleDbCommand();
                cmdSelect.Connection = con;

                cmdSelect.CommandText = _strQuery;

                // OleDbDataAdapter 객체 생성 및 Select Query 연결
                adap = new OleDbDataAdapter();
                adap.SelectCommand = cmdSelect;

                // DataTable 객체 생성
                dataTable = new DataTable("Data");

                // OleDbDataAdapter 객체의 Fill() 메쏘드를 사용하여 데이터소스 -> DataTable 객체에 채우도록 함
                adap.Fill(dataTable);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // OleDbConnection 객체 Close
                con.Close();
            }

            return dataTable;
        }

        public DataTable SelectDataQueryOld(string _db_path1, string _db_path2, string _strQuery)
        {
            //string connStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _db_path1 + ";Jet OLEDB:Database Password=";
            //con = new System.Data.OleDb.OleDbConnection();
            //con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0};Jet OLEDB:Database Password=", _db_path1.Trim());
            //con.Open();

            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}", _db_path1.Trim());
                con.Open();

                _strQuery = @"SELECT TB_TAKE.*, TB_MASTER_ATM.* FROM [" + _db_path1 + "].[TB_TAKE] AS TB_TAKE LEFT OUTER JOIN [" + _db_path2 + "].[TB_MASTER_ATM] AS TB_MASTER_ATM ON TB_MASTER_ATM.ATM_ID = TB_TAKE.ATM_ID";

                // OleDbCommand 객체 생성
                OleDbCommand cmdSelect = new OleDbCommand();
                cmdSelect.Connection = con;

                cmdSelect.CommandText = _strQuery;

                // OleDbDataAdapter 객체 생성 및 Select Query 연결
                adap = new OleDbDataAdapter();
                adap.SelectCommand = cmdSelect;

                // DataTable 객체 생성
                dataTable = new DataTable("Data");

                // OleDbDataAdapter 객체의 Fill() 메쏘드를 사용하여 데이터소스 -> DataTable 객체에 채우도록 함
                adap.Fill(dataTable);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // OleDbConnection 객체 Close
                con.Close();
            }

            return dataTable;
        }

        /// <summary>
        /// Access DB Hogi 데이터 가져오기
        /// </summary>
        /// <param name="_strQuery"></param>
        /// <param name="_intColumnCount"></param>
        /// <returns></returns>
        public string SelectHogiOleDb(string _strAccessDBFileName, string _strQuery)
        {
            string strResult = "";

            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}", _strAccessDBFileName.Trim());
                con.Open();

                // OleDbCommand 객체 생성
                OleDbCommand cmdSelect = new OleDbCommand();
                cmdSelect.Connection = con;

                cmdSelect.CommandText = _strQuery;

                // OleDbDataAdapter 객체 생성 및 Select Query 연결
                adap = new OleDbDataAdapter();
                adap.SelectCommand = cmdSelect;

                // DataTable 객체 생성
                dataTable = new DataTable("Data");

                // OleDbDataAdapter 객체의 Fill() 메쏘드를 사용하여 데이터소스 -> DataTable 객체에 채우도록 함
                adap.Fill(dataTable);

                strResult = dataTable == null || dataTable.Rows.Count <= 0 ? "0" : dataTable.Rows[0]["ATM_NAME"].ToString().Trim(); 
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // OleDbConnection 객체 Close
                con.Close();
            }

            return strResult;
        }

        /// <summary>
        /// Access DB Coil Name 데이터 가져오기
        /// </summary>
        /// <param name="_strQuery"></param>
        /// <param name="_intColumnCount"></param>
        /// <returns></returns>
        public string SelectCoilNameOleDb(string _strAccessDBFileName, string _strQuery)
        {
            string strResult = "";

            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}", _strAccessDBFileName.Trim());
                con.Open();

                // OleDbCommand 객체 생성
                OleDbCommand cmdSelect = new OleDbCommand();
                cmdSelect.Connection = con;

                cmdSelect.CommandText = _strQuery;

                // OleDbDataAdapter 객체 생성 및 Select Query 연결
                adap = new OleDbDataAdapter();
                adap.SelectCommand = cmdSelect;

                // DataTable 객체 생성
                dataTable = new DataTable("Data");

                // OleDbDataAdapter 객체의 Fill() 메쏘드를 사용하여 데이터소스 -> DataTable 객체에 채우도록 함
                adap.Fill(dataTable);

                strResult = dataTable == null || dataTable.Rows.Count <= 0 ? "" : dataTable.Rows[0]["COIL_CODE"].ToString().Trim().Replace("권선", "");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // OleDbConnection 객체 Close
                con.Close();
            }

            return strResult;
        }

        /// <summary>
        /// Access DB Control Name 데이터 가져오기
        /// </summary>
        /// <param name="_strQuery"></param>
        /// <param name="_intColumnCount"></param>
        /// <returns></returns>
        public string SelectControlNameOleDb(string _strAccessDBFileName, string _strQuery)
        {
            string strResult = "";

            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}", _strAccessDBFileName.Trim());
                con.Open();

                // OleDbCommand 객체 생성
                OleDbCommand cmdSelect = new OleDbCommand();
                cmdSelect.Connection = con;

                cmdSelect.CommandText = _strQuery;

                // OleDbDataAdapter 객체 생성 및 Select Query 연결
                adap = new OleDbDataAdapter();
                adap.SelectCommand = cmdSelect;

                // DataTable 객체 생성
                dataTable = new DataTable("Data");

                // OleDbDataAdapter 객체의 Fill() 메쏘드를 사용하여 데이터소스 -> DataTable 객체에 채우도록 함
                adap.Fill(dataTable);

                strResult = dataTable == null || dataTable.Rows.Count <= 0 ? "" : dataTable.Rows[0]["CTRL_CODE"].ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // OleDbConnection 객체 Close
                con.Close();
            }

            return strResult;
        }

        /// <summary>
        /// Access DB Power Cabinet 데이터 가져오기
        /// </summary>
        /// <param name="_strQuery"></param>
        /// <param name="_intColumnCount"></param>
        /// <returns></returns>
        public string SelectPowerCabinetOleDb(string _strAccessDBFileName, string _strQuery)
        {
            string strResult = "";

            try
            {
                // OleDbConnection 객체 생성 및 Open
                con = new OleDbConnection();
                con.ConnectionString = string.Format("Provider=Microsoft.JET.OLEDB.4.0; data source={0}", _strAccessDBFileName.Trim());
                con.Open();

                // OleDbCommand 객체 생성
                OleDbCommand cmdSelect = new OleDbCommand();
                cmdSelect.Connection = con;

                cmdSelect.CommandText = _strQuery;

                // OleDbDataAdapter 객체 생성 및 Select Query 연결
                adap = new OleDbDataAdapter();
                adap.SelectCommand = cmdSelect;

                // DataTable 객체 생성
                dataTable = new DataTable("Data");

                // OleDbDataAdapter 객체의 Fill() 메쏘드를 사용하여 데이터소스 -> DataTable 객체에 채우도록 함
                adap.Fill(dataTable);

                strResult = dataTable == null || dataTable.Rows.Count <= 0 ? "" : dataTable.Rows[0]["GRP_NAME"].ToString().Trim();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
            }
            finally
            {
                // OleDbConnection 객체 Close
                con.Close();
            }

            return strResult;
        }
    }
}
