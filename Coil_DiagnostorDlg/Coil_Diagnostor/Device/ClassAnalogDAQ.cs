using NationalInstruments.DAQmx;
using System;
using System.Data;
using System.Windows.Forms;

namespace Coil_Diagnostor.Device
{
	public class ClassAnalogDAQProcess
    {
		// Task
        protected Task myTask = null;
        private DataTable dataTable = null;
		
		// 호출 Form 설정
		//frmRelayTester frm;
		
        public ClassAnalogDAQProcess()
        {
        }

        #region ◈ AI Channel Reader Sample
		
        /// <summary>
        /// AI Channel One Reader 
        /// </summary>
        /// <param name="channelName"></param>
        public bool startAITaskOneReader(string _strChannelName, double _dMinimumValueNumeric, double _dMaximumValueNumeric, ref double _dMeasurementValue)
        {
            bool boolResult = false;
            dataTable = new DataTable();

            try
            {
                //Create a new task
                using (myTask = new Task())
                {
                    //Create a virtual channel
                    myTask.AIChannels.CreateVoltageChannel(_strChannelName.Trim(), "", AITerminalConfiguration.Rse
                        , _dMinimumValueNumeric, _dMaximumValueNumeric, AIVoltageUnits.Volts);

                    AnalogMultiChannelReader reader = new AnalogMultiChannelReader(myTask.Stream);

                    //Verify the Task
                    myTask.Control(TaskAction.Verify);

                    //Plot Multiple Channels to the table
                    double[] data = reader.ReadSingleSample();
                    _dMeasurementValue = data[0];

                    boolResult = true;
                }
            }
            catch (DaqException exception)
            {
                boolResult = false;
                MessageBox.Show(exception.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// AI Channel Self Test Reader 
        /// </summary>
        /// <param name="channelName"></param>
        public bool startAITaskSelfTestReader(string _strChannelName, double _dMinimumValueNumeric, double _dMaximumValueNumeric, ref double _dMeasurementValue
            , double dRate, int intSamplesPerChannel)
        {
            bool boolResult = false;
            dataTable = new DataTable();

            try
            {
                //Create a new task
                using (myTask = new Task())
                {
                    //Create a virtual channel
                    myTask.AIChannels.CreateVoltageChannel(_strChannelName.Trim(), "", AITerminalConfiguration.Rse
                        , _dMinimumValueNumeric, _dMaximumValueNumeric, AIVoltageUnits.Volts);

                    // Configure the timing parameters
                    myTask.Timing.ConfigureSampleClock("", dRate,
                        SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, intSamplesPerChannel);

                    AnalogMultiChannelReader reader = new AnalogMultiChannelReader(myTask.Stream);

                    //Verify the Task
                    myTask.Control(TaskAction.Verify);

                    //Plot Multiple Channels to the table
                    double[,] data = reader.ReadMultiSample(intSamplesPerChannel);
                    for (int i = 0; i < data.Length; i++)
                    {
                        if (data[0, i] > 8d)
                        {
                            boolResult = true;
                            break;
                        }
                    }
                }
            }
            catch (DaqException exception)
            {
                boolResult = false;
                MessageBox.Show(exception.Message);
            }

            return boolResult;
        }

        /// <summary>
        /// AI Channel Multi Reader 
        /// </summary>
        /// <param name="channelName"></param>
        public bool startAITaskMultiReader(string _strChannelName, double _dMinimumValueNumeric, double _dMaximumValueNumeric, ref double _dMeasurementValue
            , double dRate, int intSamplesPerChannel)
        {
            bool boolResult = false;
            dataTable = new DataTable();

            try
            {
                //Create a new task
                using (myTask = new Task())
                {
                    //Create a virtual channel
                    myTask.AIChannels.CreateVoltageChannel(_strChannelName.Trim(), "", AITerminalConfiguration.Rse
                        , _dMinimumValueNumeric, _dMaximumValueNumeric, AIVoltageUnits.Volts);

                    // Configure the timing parameters
                    myTask.Timing.ConfigureSampleClock("", dRate,
                        SampleClockActiveEdge.Rising, SampleQuantityMode.ContinuousSamples, intSamplesPerChannel);

                    AnalogMultiChannelReader reader = new AnalogMultiChannelReader(myTask.Stream);

                    //Verify the Task
                    myTask.Control(TaskAction.Verify);

                    //Plot Multiple Channels to the table
                    double[,] data = reader.ReadMultiSample(intSamplesPerChannel);

                    double dSum = 0;
                    int intCount = 0, intDataCount = 0;

                    intDataCount = Convert.ToInt32(Convert.ToDouble(intSamplesPerChannel) * (30d / 100d));

                    double[] arrayData = new double[intSamplesPerChannel];

                    for (int i = 0; i < data.Length; i++)
                    {
                        arrayData[i] = data[0, i];
                    }

                    Array.Sort(arrayData);

                    for (int i = 30; i < arrayData.Length - 30; i++)
                    {
                        dSum += arrayData[i];
                        intCount++;
                    }

                    _dMeasurementValue = dSum / Convert.ToDouble(intCount);

                    boolResult = true;
                }
            }
            catch (DaqException exception)
            {
                boolResult = false;
                MessageBox.Show(exception.Message);
            }

            return boolResult;
        }

		#endregion
	}
}
