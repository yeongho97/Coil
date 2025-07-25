using System;
using NationalInstruments.DAQmx;
using System.Threading;

namespace Coil_Diagnostor.Device
{
	public class ClassDigitalDAQProcess
	{
        public Task myTask = null;
        public Task mainTask = null;
        public bool m_bConnected = false;
        public NationalInstruments.DAQmx.Device dev;

		/// <summary>
		/// DigitalDAQ 설정
		/// </summary>
		/// <param name="strAddress">주소</param>
		/// <returns></returns>
        public bool DigitalDAQ_Setting(string _strDigitalDaq)
        {
            try
            {
                dev = DaqSystem.Local.LoadDevice(_strDigitalDaq);
                dev.SelfTest();
                //dev.Dispose();

				m_bConnected = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);

				m_bConnected = false;
                return m_bConnected;
            }

            return m_bConnected;
        }

		/// <summary>
		/// DigitalDAQ 연결 상태
		/// </summary>
		/// <returns></returns>
        public bool IsConnected()
        {
            return m_bConnected;
        }

        /// <summary>
        /// DigitalDAQ Channel ON / OFF Write
        /// </summary>
        /// <param name="_strChannel">Channel 값</param>
        /// <param name="_bVal">bool 값</param>
        /// <returns></returns>
        public bool MainDigitalDAQ_WriteChannel(string _strChannel, bool _bVal)
        {
            try
            {
                if (mainTask != null) mainTask.Dispose();

                mainTask = new Task();
                mainTask.DOChannels.CreateChannel(_strChannel, "", ChannelLineGrouping.OneChannelForEachLine);
                bool[] dataArray = new bool[mainTask.DOChannels.Count];

                for (int i = 0; i < mainTask.DOChannels.Count; i++)
                    dataArray[i] = _bVal;

                DigitalMultiChannelWriter writer = new DigitalMultiChannelWriter(mainTask.Stream);
                writer.WriteSingleSampleSingleLine(true, dataArray);
                mainTask.Dispose();
                return true;
            }
            catch (DaqException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
        }

		/// <summary>
		/// DigitalDAQ Channel ON / OFF Write
		/// </summary>
		/// <param name="_strChannel">Channel 값</param>
		/// <param name="_bVal">bool 값</param>
		/// <returns></returns>
        public bool DigitalDAQ_WriteChannel(string _strChannel, bool _bVal)
        {
            try
            {
                if (myTask != null) myTask.Dispose();

                myTask = new Task();
                myTask.DOChannels.CreateChannel(_strChannel, "", ChannelLineGrouping.OneChannelForEachLine);
                bool[] dataArray = new bool[myTask.DOChannels.Count];
                
				for (int i = 0; i < myTask.DOChannels.Count; i++)
                    dataArray[i] = _bVal;
                
				DigitalMultiChannelWriter writer = new DigitalMultiChannelWriter(myTask.Stream);
                writer.WriteSingleSampleSingleLine(true, dataArray);
                myTask.Dispose();
                return true;
            }
            catch (DaqException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// DigitalDAQ Channel ON / OFF Reader
        /// </summary>
        /// <param name="_strChannel">Channel 값</param>
        /// <param name="_bVal">bool 값</param>
        /// <returns></returns>
        public bool DigitalDAQ_ReaderChannel(string _strChannel)
        {
            bool boolData = false;

            try
            {
                if (myTask != null) myTask.Dispose();
                
                bool[] readData;

                myTask = new Task();
                myTask.DIChannels.CreateChannel(_strChannel, "", ChannelLineGrouping.OneChannelForEachLine);

                DigitalSingleChannelReader myDigitalReader = new DigitalSingleChannelReader(myTask.Stream);

                //Read the digital channel
                for (int i = 0; i < 10; i++)
                {
                    readData = myDigitalReader.ReadSingleSampleMultiLine();

                    if (readData[0])
                    {
                        boolData = readData[0];
                        break;
                    }

                    Thread.Sleep(100);
                }

                myTask.Dispose();
                return boolData;
            }
            catch (DaqException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return boolData;
            }
        }

		/// <summary>
		/// DigitalDAQ Channel Close
		/// </summary>
		/// <returns></returns>
        public bool DigitalDAQ_CloseChannel()
        {
            try
            {
                if (myTask == null) return false;

                myTask.Stop();
                bool[] dataArray = new bool[myTask.DOChannels.Count];

                for (int i = 0; i < myTask.DOChannels.Count; i++)
                    dataArray[i] = false;

                DigitalMultiChannelWriter writer = new DigitalMultiChannelWriter(myTask.Stream);
                writer.WriteSingleSampleSingleLine(true, dataArray);

                myTask.Start();
                System.Threading.Thread.Sleep(500);

                myTask.Dispose();
                myTask = null;
                return true;
            }
            catch (DaqException ex)
            {
                System.Diagnostics.Debug.Print(ex.Message);
                return false;
            }
        }
	}
}
