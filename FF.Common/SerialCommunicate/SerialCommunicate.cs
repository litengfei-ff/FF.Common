#if NET461
using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks; 

namespace FF.Common.SerialCommunicate
{
    public class SerialCommunicate
    {
        private SerialPort sp;
        public SerialCommunicate()
        {
            sp = new SerialPort();
            sp.ErrorReceived += Sp_ErrorReceived;
        }

        private void Sp_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            throw new NotImplementedException();
        }

        public void Open(string portName)
        {
            sp.PortName = portName;
            if (!sp.IsOpen) sp.Open();
        }

        public void Close()
        {
            sp.Close();
        }

        /// <summary>
        /// 获取串口名数组
        /// </summary>
        /// <returns></returns>
        public string[] GetPortNames()
        {
            return SerialPort.GetPortNames();
        }

        /// <summary>
        /// 设置参数
        /// </summary>
        /// <param name="baudRate">波特率</param>
        /// <param name="dataBits">数据位长度</param>
        /// <param name="readTimeOut">读取超时时间</param>
        /// <param name="writeTimeOut">写入超时时间</param>
        /// <param name="parity">奇偶检验</param>
        /// <param name="stopBits">停止位</param>
        /// <param name="handshake">握手协议</param>
        public void Set(int baudRate, int dataBits, int readTimeOut, int writeTimeOut, Parity parity, StopBits stopBits, Handshake handshake)
        {
            sp.BaudRate = baudRate;
            sp.DataBits = dataBits;
            sp.ReadTimeout = readTimeOut;
            sp.WriteTimeout = writeTimeOut;
            sp.Parity = parity;
            sp.StopBits = stopBits;
            sp.Handshake = handshake;
        }

        public void Send(string msg)
        {
            sp.WriteLine(msg);
        }

        public void SetReceiveHandler(Action<object, SerialDataReceivedEventArgs> receiveHandler)
        {
            sp.DataReceived += (o, sre) => { receiveHandler(o, sre); };
        }

    }
}
#endif