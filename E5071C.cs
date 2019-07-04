using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Mitec;
using StateFile;
using MitecCommon;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Globalization;

namespace NetworkAnalyzer
{
    public class E5071C : INetworkAnalyzer
    {
        private List<ChanInfor> ChanInfo;
        private VISAConnect NaConnect;
        private string NaName;
        private ReWriteCommand ReWriteNaConnect;
        private Dictionary<Int32, string> ChanLookUp;
        private Dictionary<string, Int32> CalKitLookUp;
        private Dictionary<string, string> CalMethodLookUp;
        private ApplicationPathInfor ActualPath = new ApplicationPathInfor();

        //初始化ActiveChNum与ActiveTrNum
        private E5071C(VISAConnect equipConnect, Int32 chNum, Int32 trNum)
        {
            this.NaConnect = equipConnect;
            //初始化Chan与CalKit查询表
            IniChanLookUpTable();
            IniCalKitLookUpTable();
            IniCalMethodLookUpTable();
        }

        public E5071C(VISAConnect equipConnect, string naName, ApplicationPathInfor pathInfor)
            : this(equipConnect, 1, 1)
        {
            this.NaName = naName;
            this.ActualPath = pathInfor;
        }

        //E5071c 没有Numbers of Channel的选项，只可以 Allocate
        private void IniChanLookUpTable()
        {
            ChanLookUp = new Dictionary<Int32, string>();
            ChanLookUp.Add(1, "D1");
            ChanLookUp.Add(2, "D1_2");
            ChanLookUp.Add(3, "D1_2_3");
            ChanLookUp.Add(4, "D12_34");
            ChanLookUp.Add(6, "D123_456");
            ChanLookUp.Add(8, "D1234_5678");
            ChanLookUp.Add(9, "D123_456_789");
            ChanLookUp.Add(12, "D123__ABC");
            ChanLookUp.Add(16, "D1234__CDEF");
        }

        private void IniCalKitLookUpTable()
        {
            CalKitLookUp = new Dictionary<string, int>();
            CalKitLookUp.Add("85033E", 1);
            CalKitLookUp.Add("85033D", 2);
            CalKitLookUp.Add("85052D", 3);
            CalKitLookUp.Add("85032F", 4);
            CalKitLookUp.Add("85032B/E", 5);
            CalKitLookUp.Add("85036B/E", 6);
            CalKitLookUp.Add("85031B", 7);
            //注意这种写法是为兼容N5230C
            CalKitLookUp.Add("85050C", 8);
            CalKitLookUp.Add("85050D", 9);
            CalKitLookUp.Add("85052C", 10);

            //TODO:

        }

        private void IniCalMethodLookUpTable()
        {
            CalMethodLookUp = new Dictionary<string, string>();
            CalMethodLookUp.Add("Solt2PortCal", "SOLT");
            CalMethodLookUp.Add("TRL2PortCal", "TRL");
            CalMethodLookUp.Add("TRANCal", "THRU");
        }

        public bool SetStateSetting(NetworkAnalyzerState networkAnalyzerState)
        {
            try
            {
                //清空error
                this.ClearError();

                this.Prest();

                //分配通道
                AllocateChannels(networkAnalyzerState.ChanCount);

                this.ChanInfo = networkAnalyzerState.ChanInfo;

                foreach (ChanInfor curChan in this.ChanInfo)
                {
                    //选择通道
                    SelectChannel(curChan.Chan);
                    //设置显示
                    SetDisplayInfor(curChan.Chan, curChan.Display);

                    //设置平均
                    SetAverageInfor(curChan.Chan, curChan.Averaging);

                    //设置激励
                    SetStimulusInfor(curChan.Chan, curChan.Stimulus);

                    //设置扫描
                    SetSweep(curChan.Chan, curChan.Sweep);

                    //分配Trace给指定Channel
                    SetNumberOfTraces(curChan.Chan, curChan.TraceCount);

                    foreach (TraceInfor curTrace in curChan.Traces)
                    {
                        //设置迹线信息
                        SetTraceInfor(curChan.Chan, curTrace);
                    }

                }

                MonitorSystem();
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("设置状态文件出错: " + Environment.NewLine + ex.Message);
            }

        }

        public bool CheckStateSetting(NetworkAnalyzerState networkAnalyzerState)
        {
            try
            {
                this.ClearError();
                CheckAllocatedChannels(networkAnalyzerState.ChanCount);

                foreach (ChanInfor curChan in networkAnalyzerState.ChanInfo)
                {

                    //选择通道
                    SelectChannel(curChan.Chan);

                    CheckDisplayInfor(curChan.Chan, curChan.Display);

                    CheckAverageInfor(curChan.Chan, curChan.Averaging);

                    CheckStimulusInfor(curChan.Chan, curChan.Stimulus);

                    CheckSweep(curChan.Chan, curChan.Sweep);

                    CheckNumberOfTraces(curChan.Chan, curChan.TraceCount);

                    foreach (TraceInfor curTrace in curChan.Traces)
                    {
                        //Check迹线信息
                        CheckTraceInfor(curChan.Chan, curTrace);
                    }

                }
                MonitorSystem();
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("验证状态文件出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Prest()
        {
            try
            {
                this.ReWriteNaConnect.Write(@":SYST:PRES");
            }
            catch (Exception ex)
            {
                throw new Exception("Preset出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void AllocateChannels(Int32 chCount)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":DISP:SPL " + ChanLookUp[chCount]);
                CheckAllocatedChannels(chCount);
            }
            catch (Exception ex)
            {
                throw new Exception("分配Channel出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckAllocatedChannels(int chCount)
        {
            try
            {
                int getchCount = 0;
                this.NaConnect.sendCmd(@":DISP:SPL?");
                string disp = this.NaConnect.readStr().Trim();
                List<int> num = (from ChanNum in ChanLookUp
                                 where ChanNum.Value == disp
                                 select ChanNum.Key).ToList();
                getchCount = num[0];

                if (getchCount != chCount)
                {
                    throw new Exception("通道数量错误: 指标为 " + chCount + " ;仪器读数为 " + getchCount);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("通道数量异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SelectChannel(Int32 chNum)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":ACT");
            }
            catch (Exception ex)
            {
                throw new Exception("选择Channel出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetDisplayInfor(int chNum, DisplayInfor displayInfor)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":TITL:DATA \"" + displayInfor.TitleLabel + "\"");
                //默认打开显示Title功能
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":TITL:STAT 1");
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置DisplayInfor出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckDisplayInfor(int chNum, DisplayInfor displayInfor)
        {
            try
            {
                DisplayInfor Disp = new DisplayInfor();
                double titleEnable = 0;
                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":TITL:DATA?");
                Disp.TitleLabel = this.NaConnect.readStr().Trim();
                if (Disp.TitleLabel != "\"" + displayInfor.TitleLabel + "\"")
                {
                    throw new Exception("通道 " + chNum + " TitleLabel错误: 指标为 " + displayInfor.TitleLabel + " ;仪器读数为 " + Disp.TitleLabel);
                }
                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":TITL?");
                titleEnable = this.NaConnect.readNum();
                if (titleEnable != 1)
                {
                    throw new Exception("通道 " + chNum + " 标题显示错误: 默认为 ON" + " ;仪器读数为 " + (IsEnable)titleEnable);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("DisplayInfor异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetAverageInfor(int chNum, AverageInfor averageInfor)
        {
            try
            {
                IncludeFactorAvgInfor inFacAvgInfor = null;
                ExcludeFactorAvgInfor exFacAvgInfor = null;
                switch (averageInfor.GetType().Name.ToString())
                {

                    case "IncludeFactorAvgInfor":
                        inFacAvgInfor = (IncludeFactorAvgInfor)averageInfor;
                        this.ReWriteNaConnect.Write(@":SENS" + chNum + ":AVER:STAT " + inFacAvgInfor.AveragingEnable);
                        this.ReWriteNaConnect.Write(@":SENS" + chNum + ":AVER:COUN " + inFacAvgInfor.AvgFactor);
                        this.ReWriteNaConnect.Write(@":TRIG:AVER " + IsEnable.ON);
                        this.ReWriteNaConnect.Write(@":SENS" + chNum + ":BAND " + inFacAvgInfor.IFBandWidth);
                        break;
                    case "ExcludeFactorAvgInfor":
                        exFacAvgInfor = (ExcludeFactorAvgInfor)averageInfor;
                        this.ReWriteNaConnect.Write(@":SENS" + chNum + ":BAND " + exFacAvgInfor.IFBandWidth);
                        break;
                    default:
                        throw new Exception("Averaging 未找到相应指标设置！");

                }
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置AverageInfor出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckAverageInfor(int chNum, AverageInfor averageInfor)
        {
            try
            {
                IncludeFactorAvgInfor inFacAvgInfor = null;
                ExcludeFactorAvgInfor exFacAvgInfor = null;
                int avgenable;
                int avgfactor;
                int avgtrigble;
                double ifbandwid;

                switch (averageInfor.GetType().Name.ToString())
                {

                    case "IncludeFactorAvgInfor":
                        inFacAvgInfor = (IncludeFactorAvgInfor)averageInfor;
                        this.NaConnect.sendCmd(@":SENS" + chNum + ":AVER:STAT?");
                        avgenable = Convert.ToInt32(this.NaConnect.readNum());
                        if ((IsEnable)avgenable != inFacAvgInfor.AveragingEnable)
                        {
                            throw new Exception("通道 " + chNum + " AveragingEnable错误: 指标为 " + inFacAvgInfor.AveragingEnable + " ;仪器读数为 " + (IsEnable)avgenable);
                        }

                        this.NaConnect.sendCmd(@":SENS" + chNum + ":AVER:COUN?");
                        avgfactor = Convert.ToInt32(this.NaConnect.readNum());
                        if (avgfactor != inFacAvgInfor.AvgFactor)
                        {
                            throw new Exception("通道 " + chNum + " AvgFactor错误: 指标为 " + inFacAvgInfor.AvgFactor + " ;仪器读数为 " + avgfactor);
                        }
                        this.NaConnect.sendCmd(@":TRIG:AVER?");
                        avgtrigble = Convert.ToInt32(this.NaConnect.readNum());
                        if ((IsEnable)avgtrigble != IsEnable.ON)
                        {
                            throw new Exception("通道 " + chNum + " Avg Trigger错误：默认为 ON " + ";仪器读数为 " + (IsEnable)avgtrigble);
                        }

                        this.NaConnect.sendCmd(@":SENS" + chNum + ":BAND?");
                        ifbandwid = this.NaConnect.readNum();
                        if (ifbandwid != inFacAvgInfor.IFBandWidth)
                        {
                            throw new Exception("通道 " + chNum + " IFBW错误: 指标为 " + inFacAvgInfor.IFBandWidth + " ;仪器读数为 " + ifbandwid);
                        }
                        break;
                    case "ExcludeFactorAvgInfor":
                        exFacAvgInfor = (ExcludeFactorAvgInfor)averageInfor;

                        this.NaConnect.sendCmd(@":SENS" + chNum + ":AVER:STAT?");
                        avgenable = Convert.ToInt32(this.NaConnect.readNum());
                        if ((IsEnable)avgenable != IsEnable.OFF)
                        {
                            throw new Exception("通道 " + chNum + " AveragingEnable错误: 默认为 OFF " + ";仪器读数为 " + (IsEnable)avgenable);
                        }

                        this.NaConnect.sendCmd(@":SENS" + chNum + ":BAND?");
                        ifbandwid = this.NaConnect.readNum();
                        if (ifbandwid != exFacAvgInfor.IFBandWidth)
                        {
                            throw new Exception("通道 " + chNum + " IFBW错误: 指标为 " + exFacAvgInfor.IFBandWidth + " ;仪器读数为 " + ifbandwid);
                        }
                        break;

                    case "IncludeSmoAperture":





                        break;
                    default:
                        throw new Exception("Averaging 未找到相应指标设置！");
                }

            }
            catch (Exception ex)
            {
                throw new Exception("AverageInfor异常: " + Environment.NewLine + ex.Message);
            }
        }

        //Stimulus 单位 MHz
        public void SetStimulusInfor(int chNum, StimulusInfor stimulusInfor)
        {
            try
            {
                StimulusInforOfFreq stisOfFreq = null;
                StimulusInforOfPower stisOfPower = null;
                switch (stimulusInfor.GetType().Name.ToString())
                {
                    case "StimulusInforOfFreq":
                        stisOfFreq = (StimulusInforOfFreq)stimulusInfor;
                        this.ReWriteNaConnect.Write(@":SENS" + chNum + ":FREQ:STAR " + stisOfFreq.Start * 1000000);
                        this.ReWriteNaConnect.Write(@":SENS" + chNum + ":FREQ:STOP " + stisOfFreq.Stop * 1000000);
                        break;

                    case "StimulusInforOfPower":
                        stisOfPower = (StimulusInforOfPower)stimulusInfor;
                        this.ReWriteNaConnect.Write(@":SOUR" + chNum + ":POW:STAR " + stisOfPower.Start);
                        this.ReWriteNaConnect.Write(@":SOUR" + chNum + ":POW:STOP " + stisOfPower.Stop);
                        break;

                    default:
                        throw new Exception("Stimulus 未找到相应指标设置！");
                }
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置StimulusInfor出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckStimulusInfor(int chNum, StimulusInfor stimulusInfor)
        {
            try
            {
                StimulusInforOfFreq stisOfFreq = null;
                StimulusInforOfPower stisOfPower = null;

                double Start;
                double Stop;

                switch (stimulusInfor.GetType().Name.ToString())
                {
                    case "StimulusInforOfFreq":
                        stisOfFreq = (StimulusInforOfFreq)stimulusInfor;
                        this.NaConnect.sendCmd(@":SENS" + chNum + ":FREQ:STAR?");
                        Start = this.NaConnect.readNum() / 1000000;
                        if (Start != stisOfFreq.Start)
                        {
                            throw new Exception("通道 " + chNum + " 起始频率错误: 指标为 " + stisOfFreq.Start + " ;仪器读数为 " + Start + "\n请确认产品频段是否在仪器范围中!");
                        }
                        this.NaConnect.sendCmd(@":SENS" + chNum + ":FREQ:STOP?");
                        Stop = this.NaConnect.readNum() / 1000000;
                        if (Stop != stisOfFreq.Stop)
                        {
                            throw new Exception("通道 " + chNum + " 终止频率错误: 指标为 " + stisOfFreq.Stop + " ;仪器读数为 " + Stop + "\n请确认产品频段是否在仪器范围中!");
                        }
                        break;

                    case "StimulusInforOfPower":
                        stisOfPower = (StimulusInforOfPower)stimulusInfor;
                        this.NaConnect.sendCmd(@":SOUR" + chNum + ":POW:STAR?");
                        Start = this.NaConnect.readNum();
                        if (Start != stisOfPower.Start)
                        {
                            throw new Exception("通道 " + chNum + " 起始功率错误: 指标为 " + stisOfPower.Start + " ;仪器读数为 " + Start);
                        }

                        this.NaConnect.sendCmd(@":SOUR" + chNum + ":POW:STOP?");
                        Stop = this.NaConnect.readNum();
                        if (Stop != stisOfPower.Stop)
                        {
                            throw new Exception("通道 " + chNum + " 终止功率错误: 指标为 " + stisOfPower.Stop + " ;仪器读数为 " + Stop);
                        }
                        break;

                    default:
                        throw new Exception("Stimulus 未找到相应指标设置！");
                }

            }
            catch (Exception ex)
            {
                throw new Exception("StimulusInfor异常: " + Environment.NewLine + ex.Message);
            }
        }

        //过程中默认SegmentData格式已写死
        //Freq单位 MHz
        public void SetSweep(int chNum, SweepInfor sweepInfor)
        {
            try
            {
                IncludeSegmentDataSweep inSegDatSweep = null;
                ExcludeSegmentDataSweep exSegDatSweep = null;
                StringBuilder outData = new StringBuilder();

                switch (sweepInfor.GetType().Name.ToString())
                {

                    case "IncludeSegmentDataSweep":
                        inSegDatSweep = (IncludeSegmentDataSweep)sweepInfor;

                        SetSweepType(chNum, inSegDatSweep.SweepType);
                        SetSweepSegmentDisplay(chNum, inSegDatSweep.SegmentDisplay);
                        SetSweepPoints(chNum, inSegDatSweep.Points);
                        SetSweepPower(chNum, inSegDatSweep.Power);

                        //<buf>
                        outData.Append("5");
                        //<mode>
                        outData.Append("," + "0");
                        //<ifbw>
                        outData.Append("," + "1");
                        //<pow>
                        outData.Append("," + "1");
                        //<del>
                        outData.Append("," + "0");
                        //<time>
                        outData.Append("," + "0");
                        //<segm>
                        outData.Append("," + inSegDatSweep.SegmentData.Count().ToString());

                        foreach (SegmentDataInfor curSeg in inSegDatSweep.SegmentData)
                        {
                            outData.Append("," + (curSeg.StartFreq * 1000000).ToString());
                            outData.Append("," + (curSeg.StopFreq * 1000000).ToString());
                            outData.Append("," + curSeg.Points.ToString());
                            outData.Append("," + curSeg.IFBW.ToString());
                            outData.Append("," + curSeg.Power.ToString());
                        }
                        ReWriteNaConnect.Write(@":SENS" + chNum + ":SEGM:DATA " + outData.ToString());

                        break;

                    case "ExcludeSegmentDataSweep":
                        exSegDatSweep = (ExcludeSegmentDataSweep)sweepInfor;

                        SetSweepType(chNum, exSegDatSweep.SweepType);
                        SetSweepSegmentDisplay(chNum, exSegDatSweep.SegmentDisplay);
                        SetSweepPoints(chNum, exSegDatSweep.Points);
                        SetSweepPower(chNum, exSegDatSweep.Power);

                        break;

                    default:
                        throw new Exception("Sweep 未找到相应指标设置！");
                }

                //连续初始化模式关闭|Hold
                ReWriteNaConnect.Write(@":INIT" + chNum + ":CONT OFF");
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置Sweep出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckSweep(int chNum, SweepInfor sweepInfor)
        {
            try
            {
                IncludeSegmentDataSweep inSegDatSweep = null;
                ExcludeSegmentDataSweep exSegDatSweep = null;
                StringBuilder outData = new StringBuilder();

                List<string> Segdatas = new List<string>();
                List<string> tempseg = new List<string>();
                List<string> seg = new List<string>();
                int segCount = 0;

                switch (sweepInfor.GetType().Name.ToString())
                {
                    case "IncludeSegmentDataSweep":
                        inSegDatSweep = (IncludeSegmentDataSweep)sweepInfor;

                        CheckSweepType(chNum, inSegDatSweep.SweepType);
                        CheckSweepSegmentDisplay(chNum, inSegDatSweep.SegmentDisplay);
                        CheckSweepPoints(chNum, inSegDatSweep.Points);
                        CheckSweepPower(chNum, inSegDatSweep.Power);

                        this.NaConnect.sendCmd(@":SENS" + chNum + ":SEGM:DATA?");
                        Segdatas = this.NaConnect.readStr().Trim().Split(',').ToList<string>();


                        tempseg = Segdatas.GetRange(0, 7);
                        if ((IsEnable)(int)Convert.ToDouble(tempseg[1]) != IsEnable.OFF)
                        {
                            throw new Exception("通道 " + chNum + " Segment FreqMode错误: 默认为 Start/Stop" + " ;仪器读数为 Center/Span");
                        }

                        if ((IsEnable)(int)Convert.ToDouble(tempseg[2]) != IsEnable.ON)
                        {
                            throw new Exception("通道 " + chNum + " Segment ListIFBW错误: 默认为 ON" + " ;仪器读数为 " + (IsEnable)Convert.ToInt32(tempseg[2]));
                        }

                        if ((IsEnable)(int)Convert.ToDouble(tempseg[3]) != IsEnable.ON)
                        {
                            throw new Exception("通道 " + chNum + " Segment ListPower错误: 默认为 ON" + " ;仪器读数为 " + (IsEnable)Convert.ToInt32(tempseg[3]));
                        }

                        if ((IsEnable)(int)Convert.ToDouble(tempseg[4]) != IsEnable.OFF)
                        {
                            throw new Exception("通道 " + chNum + " Segment ListDelay错误: 默认为 OFF" + " ;仪器读数为 " + (IsEnable)Convert.ToInt32(tempseg[4]));
                        }

                        if ((IsEnable)(int)Convert.ToDouble(tempseg[5]) != IsEnable.OFF)
                        {
                            throw new Exception("通道 " + chNum + " Segment ListTime错误: 默认为 OFF" + " ;仪器读数为 " + (IsEnable)Convert.ToInt32(tempseg[5]));
                        }


                        segCount = (int)Convert.ToDouble(tempseg[6]);

                        tempseg = Segdatas.GetRange(7, Segdatas.Count() - 7);
                        if (segCount != tempseg.Count / 5)
                        {
                            throw new Exception("通道 " + chNum + " SegmentData数组错误: 指标为 " + inSegDatSweep.SegmentData.Count() + " ;仪器读数为 " + segCount);
                        }


                        for (int count = 0; count < segCount; count++)
                        {
                            seg = tempseg.GetRange(count * 5, 5);


                            if (Convert.ToDouble(seg[0]) / 1000000 != inSegDatSweep.SegmentData[count].StartFreq)
                            {
                                throw new Exception("通道 " + chNum + " SegmentData 第 " + (count + 1) + " 组StartFreq错误: 指标为 " + inSegDatSweep.SegmentData[count].StartFreq + " ;仪器读数为 " + (Convert.ToDouble(seg[0]) / 1000000));
                            }
                            if (Convert.ToDouble(seg[1]) / 1000000 != inSegDatSweep.SegmentData[count].StopFreq)
                            {
                                throw new Exception("通道 " + chNum + " SegmentData 第 " + (count + 1) + " 组StopFreq错误: 指标为 " + inSegDatSweep.SegmentData[count].StopFreq + " ;仪器读数为 " + (Convert.ToDouble(seg[1]) / 1000000));
                            }
                            if (Convert.ToDouble(seg[2]) != inSegDatSweep.SegmentData[count].Points)
                            {
                                throw new Exception("通道 " + chNum + " SegmentData 第 " + (count + 1) + " 组Points错误: 指标为 " + inSegDatSweep.SegmentData[count].Points + " ;仪器读数为 " + Convert.ToDouble(seg[2]));
                            }
                            if (Convert.ToDouble(seg[3]) != inSegDatSweep.SegmentData[count].IFBW)
                            {
                                throw new Exception("通道 " + chNum + " SegmentData 第 " + (count + 1) + " 组IFBW错误: 指标为 " + inSegDatSweep.SegmentData[count].IFBW + " ;仪器读数为 " + seg[3]);
                            }
                            if (Convert.ToDouble(seg[4]) != inSegDatSweep.SegmentData[count].Power)
                            {
                                throw new Exception("通道 " + chNum + " SegmentData 第 " + (count + 1) + " 组Power错误: 指标为 " + inSegDatSweep.SegmentData[count].StopFreq + " ;仪器读数为 " + seg[4]);
                            }
                        }

                        break;

                    case "ExcludeSegmentDataSweep":
                        exSegDatSweep = (ExcludeSegmentDataSweep)sweepInfor;

                        CheckSweepType(chNum, exSegDatSweep.SweepType);
                        CheckSweepSegmentDisplay(chNum, exSegDatSweep.SegmentDisplay);
                        CheckSweepPoints(chNum, exSegDatSweep.Points);
                        CheckSweepPower(chNum, exSegDatSweep.Power);

                        break;

                    default:
                        throw new Exception("Sweep 未找到相应指标设置！");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Sweep异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetNumberOfTraces(int chNum, int TraceCount)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":PAR:COUN " + TraceCount);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置迹线数量出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckNumberOfTraces(int chNum, int TraceCount)
        {
            try
            {
                int gettrcount = 0;
                this.NaConnect.sendCmd(@"CALC" + chNum + ":PAR:COUN?");
                gettrcount = Convert.ToInt16(this.NaConnect.readStr().Trim());
                if (gettrcount != TraceCount)
                {
                    throw new Exception("通道 " + chNum + " 迹线数量错误: 指标为 " + TraceCount + " ;仪器读数为 " + gettrcount);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("通道迹线数量异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetTraceInfor(int chNum, TraceInfor traceInfor)
        {
            try
            {
                TraceInforOfFinal finalTrace = null;
                TraceInforOfTune tuneTrace = null;
                switch (traceInfor.GetType().Name.ToString())
                {
                    case "TraceInforOfFinal":
                        finalTrace = (TraceInforOfFinal)traceInfor;
                        //选择迹线
                        SelectTrace(chNum, finalTrace.Trace);
                        //设置Trace format
                        SetTraceFormat(chNum, finalTrace.Format);
                        //设置Trace measurement
                        SetTraceMeas(chNum, finalTrace.Trace, finalTrace.Measurement);
                        //设置Trace Scale
                        SetTraceScale(chNum, finalTrace.Trace, finalTrace.Scale);
                        //设置Trace Smooth
                        SetTraceSmooth(chNum, finalTrace.Trace, finalTrace.Smooth);
                        break;

                    case "TraceInforOfTune":
                        tuneTrace = (TraceInforOfTune)traceInfor;
                        //选择迹线
                        SelectTrace(chNum, tuneTrace.Trace);
                        //设置Trace format
                        SetTraceFormat(chNum, tuneTrace.Format);
                        //设置Trace measurement
                        SetTraceMeas(chNum, tuneTrace.Trace, tuneTrace.Measurement);
                        //设置Trace Scale
                        SetTraceScale(chNum, tuneTrace.Trace, tuneTrace.Scale);
                        //设置Trace Smooth
                        SetTraceSmooth(chNum, tuneTrace.Trace, tuneTrace.Smooth);
                        //设置极限测试
                        SetLimitTest(chNum, tuneTrace.LimitTest);
                        //设置Marker Couple
                        SetMarkCouple(chNum, tuneTrace.MarkerCouple);


                        foreach (MarkerSetInfor curMarkset in tuneTrace.MarkerSet)
                        {
                            SetMark(chNum, curMarkset);
                        }

                        break;

                    default:
                        throw new Exception("TraceInfor未找到相应指标设置!");
                }
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置TraceInfor出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckTraceInfor(int chNum, TraceInfor traceInfor)
        {
            try
            {
                TraceInforOfFinal finalTrace = null;
                TraceInforOfTune tuneTrace = null;
                switch (traceInfor.GetType().Name.ToString())
                {
                    case "TraceInforOfFinal":
                        finalTrace = (TraceInforOfFinal)traceInfor;
                        //选择迹线
                        SelectTrace(chNum, finalTrace.Trace);
                        //Check Trace format
                        CheckTraceFormat(chNum, finalTrace.Format);
                        //Check Trace measurement
                        CheckTraceMeas(chNum, finalTrace.Trace, finalTrace.Measurement);
                        //Check Trace Scale
                        CheckTraceScale(chNum, finalTrace.Trace, finalTrace.Scale);
                        //Check Trace Smooth
                        CheckTraceSmooth(chNum, finalTrace.Trace, finalTrace.Smooth);

                        break;

                    case "TraceInforOfTune":
                        tuneTrace = (TraceInforOfTune)traceInfor;
                        //选择迹线
                        SelectTrace(chNum, tuneTrace.Trace);
                        //Check Trace format
                        CheckTraceFormat(chNum, tuneTrace.Format);
                        //Check Trace measurement
                        CheckTraceMeas(chNum, tuneTrace.Trace, tuneTrace.Measurement);
                        //Check Trace Scale
                        CheckTraceScale(chNum, tuneTrace.Trace, tuneTrace.Scale);
                        //Check Trace Smooth
                        CheckTraceSmooth(chNum, tuneTrace.Trace, tuneTrace.Smooth);
                        //Check 极限测试
                        CheckLimitTest(chNum, tuneTrace.LimitTest);
                        //Check Marker Couple
                        CheckMarkCouple(chNum, tuneTrace.MarkerCouple);

                        foreach (MarkerSetInfor curMarkset in tuneTrace.MarkerSet)
                        {
                            CheckMark(chNum, curMarkset);
                        }

                        break;

                    default:
                        throw new Exception("TraceInfor未找到相应指标设置!");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("TraceInfor异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SelectTrace(int chNum, int TrNum)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":PAR" + TrNum + ":SEL");
            }
            catch (Exception ex)
            {
                throw new Exception("选择Trace出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetTraceFormat(int chNum, Format TrFormat)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":SEL:FORM " + TrFormat);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置TraceFormat出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckTraceFormat(int chNum, Format TrFormat)
        {
            try
            {
                string trformat = "";
                this.NaConnect.sendCmd(@":CALC" + chNum + ":SEL:FORM?");
                trformat = this.NaConnect.readStr().Trim();
                if (trformat != TrFormat.ToString())
                {
                    throw new Exception("通道 " + chNum + " TraceFormat错误: 指标为 " + TrFormat.ToString() + " ;仪器读数为 " + trformat);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("TraceFormat异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetTraceMeas(int chNum, int TrNum, SPar measurement)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":PAR" + TrNum + ":DEF " + measurement);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置TraceMeas出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckTraceMeas(int chNum, int TrNum, SPar measurement)
        {
            try
            {
                string meas = "";
                this.NaConnect.sendCmd(@":CALC" + chNum + ":PAR" + TrNum + ":DEF?");
                meas = this.NaConnect.readStr().Trim();
                if (meas != measurement.ToString())
                {
                    throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " Measurement错误: 指标为 " + measurement.ToString() + " ;仪器读数为 " + meas);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("通道迹线Measurement异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetTraceScale(int chNum, int TrNum, ScaleInfor scaleInfor)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":Y:DIV " + scaleInfor.Divisions);
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":TRAC" + TrNum + ":Y:PDIV " + scaleInfor.ScaleDivision);
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":TRAC" + TrNum + ":Y:RPOS " + scaleInfor.ReferencePosition);
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":TRAC" + TrNum + ":Y:RLEV " + scaleInfor.ReferenceValue);
                MonitorSystem();
            }

            catch (Exception ex)
            {
                throw new Exception("设置TraceScale出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckTraceScale(int chNum, int TrNum, ScaleInfor scaleInfor)
        {
            try
            {
                ScaleInfor scale = new ScaleInfor();

                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":Y:DIV?");
                scale.Divisions = this.NaConnect.readNum();
                if (scale.Divisions != scaleInfor.Divisions)
                {
                    throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " Divisions错误: 指标为 " + scaleInfor.Divisions + " ;仪器读数为 " + scale.Divisions);
                }
                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":TRAC" + TrNum + ":Y:PDIV?");
                scale.ScaleDivision = this.NaConnect.readNum();
                if (scale.ScaleDivision != scaleInfor.ScaleDivision)
                {
                    throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " ScaleDivision错误: 指标为 " + scaleInfor.ScaleDivision + " ;仪器读数为 " + scale.ScaleDivision);
                }
                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":TRAC" + TrNum + ":Y:RPOS?");
                scale.ReferencePosition = this.NaConnect.readNum();
                if (scale.ReferencePosition != scaleInfor.ReferencePosition)
                {
                    throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " ReferencePosition错误: 指标为 " + scaleInfor.ReferencePosition + " ;仪器读数为 " + scale.ReferencePosition);
                }
                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":TRAC" + TrNum + ":Y:RLEV?");
                scale.ReferenceValue = this.NaConnect.readNum();
                if (scale.ReferenceValue != scaleInfor.ReferenceValue)
                {
                    throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " ReferenceValue错误: 指标为 " + scaleInfor.ReferenceValue + " ;仪器读数为 " + scale.ReferenceValue);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("TraceScale异常: " + Environment.NewLine + ex.Message);
            }
        }

        //Stimulus单位 MHz|dB
        public void SetLimitTest(int chNum, LimitTestInfor limitTestInfor)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":LIM:DISP " + limitTestInfor.LimitLineEnable);
                StringBuilder outData = new StringBuilder();
                outData.Append(limitTestInfor.LimitTable.Count().ToString());

                foreach (LimitTableInfor limitTable in limitTestInfor.LimitTable)
                {
                    outData.Append("," + ((Int32)(limitTable.Type)).ToString());
                    outData.Append("," + (limitTable.BeginStimulus * 1000000).ToString());
                    outData.Append("," + (limitTable.EndStimulus * 1000000).ToString());
                    outData.Append("," + limitTable.BeginResponse.ToString());
                    outData.Append("," + limitTable.EndResponse.ToString());
                }

                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":LIM:DATA " + outData.ToString());
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置LimitTest出错: " + Environment.NewLine + ex.Message);
            }

        }

        public void CheckLimitTest(int chNum, LimitTestInfor limitTestInfor)
        {
            try
            {
                int enlimitline = 0;
                int limCount = 0;
                List<string> Limitdatas = new List<string>();
                List<string> templim = new List<string>();

                this.NaConnect.sendCmd(@":CALC" + chNum + ":LIM:DISP?");
                enlimitline = Convert.ToInt32(this.NaConnect.readNum());

                if ((IsEnable)enlimitline != limitTestInfor.LimitLineEnable)
                {
                    throw new Exception("通道 " + chNum + " LimitLineEnable错误: 指标为 " + limitTestInfor.LimitLineEnable + " ;仪器读数为 " + (IsEnable)enlimitline);
                }


                this.NaConnect.sendCmd(@":CALC" + chNum + ":LIM:DATA?");
                Limitdatas = this.NaConnect.readStr().Trim().Split(',').ToList();
                limCount = (int)Convert.ToDouble(Limitdatas[0]);
                if (limCount != limitTestInfor.LimitTable.Count())
                {
                    throw new Exception("通道 " + chNum + " LimitTable数量错误: 指标为 " + limitTestInfor.LimitTable.Count() + " ;仪器读数为 " + limCount);
                }

                for (int count = 0; count < limCount; count++)
                {
                    templim = Limitdatas.GetRange(count * 5 + 1, 5);
                    if ((LimitType)(int)Convert.ToDouble(templim[0]) != limitTestInfor.LimitTable[count].Type)
                    {
                        throw new Exception("通道 " + chNum + " LimitTable 第 " + (count + 1) + " 组Type错误: 指标为 " + limitTestInfor.LimitTable[count].Type + " ;仪器读数为 " + (LimitType)(int)Convert.ToDouble(templim[0]));
                    }

                    if (Convert.ToDouble(templim[1]) / 1000000 != limitTestInfor.LimitTable[count].BeginStimulus)
                    {
                        throw new Exception("通道 " + chNum + " LimitTable 第 " + (count + 1) + " 组BeginStimulus错误: 指标为 " + limitTestInfor.LimitTable[count].BeginStimulus + " ;仪器读数为 " + (Convert.ToDouble(templim[1]) / 1000000));
                    }

                    if (Convert.ToDouble(templim[2]) / 1000000 != limitTestInfor.LimitTable[count].EndStimulus)
                    {
                        throw new Exception("通道 " + chNum + " LimitTable 第 " + (count + 1) + " 组EndStimulus错误: 指标为 " + limitTestInfor.LimitTable[count].EndStimulus + " ;仪器读数为 " + (Convert.ToDouble(templim[2]) / 1000000));
                    }

                    if (Convert.ToDouble(templim[3]) != limitTestInfor.LimitTable[count].BeginResponse)
                    {
                        throw new Exception("通道 " + chNum + " LimitTable 第 " + (count + 1) + " 组BeginResponse错误: 指标为 " + limitTestInfor.LimitTable[count].BeginResponse + " ;仪器读数为 " + templim[3]);
                    }

                    if (Convert.ToDouble(templim[4]) != limitTestInfor.LimitTable[count].EndResponse)
                    {
                        throw new Exception("通道 " + chNum + " LimitTable 第 " + (count + 1) + " 组BeginResponse错误: 指标为 " + limitTestInfor.LimitTable[count].EndResponse + " ;仪器读数为 " + templim[4]);
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception("LimitTest异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetMarkCouple(int chNum, IsEnable markCouple)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":MARK:COUP " + markCouple);
                MonitorSystem();

            }
            catch (Exception ex)
            {
                throw new Exception("设置MarkCouple出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckMarkCouple(int chNum, IsEnable markCouple)
        {
            try
            {
                int markcp = 0;
                this.NaConnect.sendCmd(@":CALC" + chNum + ":MARK:COUP?");
                markcp = Convert.ToInt32(this.NaConnect.readNum());
                if ((IsEnable)markcp != markCouple)
                {
                    throw new Exception("通道 " + chNum + " MarkCouple错误: 指标为 " + markCouple + " ;仪器读数为 " + (IsEnable)markcp);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("MarkCouple异常: " + Environment.NewLine + ex.Message);
            }
        }

        //单位 MHz
        public void SetMark(int chNum, MarkerSetInfor markSetInfor)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":MARK" + markSetInfor.MarkerNo + ":STAT " + markSetInfor.MarkerEnable);
                this.ReWriteNaConnect.Write(@":CALC" + chNum + ":MARK" + markSetInfor.MarkerNo + ":X " + markSetInfor.StimulusValue * 1000000);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置SetMark出错: " + Environment.NewLine + ex.Message);
            }
        }

        //单位 Hz
        public void SetSweepPower(int chNum, SweepPowerInfor sweepPowerInfor)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":SOUR" + chNum + ":POW " + sweepPowerInfor.Power);
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":FREQ " + sweepPowerInfor.CWFreq);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置SweepPower出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckSweepPower(int chNum, SweepPowerInfor sweepPowerInfor)
        {
            try
            {
                double pow = 0;
                double freq = 0;
                this.NaConnect.sendCmd(@":SOUR" + chNum + ":POW?");
                pow = this.NaConnect.readNum();
                if (pow != sweepPowerInfor.Power)
                {
                    throw new Exception("通道 " + chNum + " Power错误: 指标为 " + sweepPowerInfor.Power + " ;仪器读数为 " + pow);
                }
                this.NaConnect.sendCmd(@":SENS" + chNum + ":FREQ?");
                freq = this.NaConnect.readNum();
                if (freq != sweepPowerInfor.CWFreq)
                {
                    throw new Exception("通道 " + chNum + " CWFreq错误: 指标为 " + sweepPowerInfor.CWFreq + " ;仪器读数为 " + freq);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("SweepPower异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetSweepPoints(int chNum, double sweepPoint)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":SWE:POIN " + sweepPoint);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置SweepPoints出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckSweepPoints(int chNum, double sweepPoint)
        {
            try
            {
                double swpoint = 0;
                this.NaConnect.sendCmd(@":SENS" + chNum + ":SWE:POIN?");
                swpoint = this.NaConnect.readNum();
                if (swpoint != sweepPoint)
                {
                    throw new Exception("通道 " + chNum + " SweepPoint错误: 指标为 " + sweepPoint + " ;仪器读数为 " + sweepPoint);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("SweepPoints异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetSweepType(int chNum, SweType sweepType)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":SWE:TYPE " + sweepType);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置SweepType出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckSweepType(int chNum, SweType sweepType)
        {
            try
            {
                string swetype = "";
                this.NaConnect.sendCmd(@":SENS" + chNum + ":SWE:TYPE?");
                swetype = this.NaConnect.readStr().Trim();
                if (swetype != sweepType.ToString())
                {
                    throw new Exception("通道 " + chNum + " SweepType错误: 指标为 " + sweepType.ToString() + " ;仪器读数为 " + swetype);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("通道SweepType异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetSweepSegmentDisplay(int chNum, SegmentDisp segmentDisplay)
        {
            try
            {
                this.ReWriteNaConnect.Write(@":DISP:WIND" + chNum + ":X:SPAC " + segmentDisplay);
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("设置SegmentDisplay出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckSweepSegmentDisplay(int chNum, SegmentDisp segmentDisplay)
        {
            try
            {
                string segdisp = "";
                this.NaConnect.sendCmd(@":DISP:WIND" + chNum + ":X:SPAC?");
                segdisp = this.NaConnect.readStr().Trim();
                if (segdisp != segmentDisplay.ToString())
                {
                    throw new Exception("通道 " + chNum + " SegmentDisplay错误: 指标为 " + segmentDisplay.ToString() + " ;仪器读数为 " + segdisp);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("通道SegmentDisplay异常: " + Environment.NewLine + ex.Message);
            }
        }

        # region "测量项"

        public void SingleSweep(Int32 chNum)
        {
            try
            {
                this.NaConnect.timeout = 300000;
                this.NaConnect.sendCmd(@":DISP:SPL?");
                string disp = this.NaConnect.readStr().Trim();
                List<int> num = (from ChanNum in ChanLookUp
                                 where ChanNum.Value == disp
                                 select ChanNum.Key).ToList();
                int AllNum = num[0];

                this.ReWriteNaConnect.Write(@":TRIG:SOUR bus");
                //防止当前通道被员工修改
                HoldAllChannels(AllNum);
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":AVER:CLE");
                this.ReWriteNaConnect.Write(@":INIT" + chNum);
                this.ReWriteNaConnect.Write(@":TRIG:SING");
                this.NaConnect.timeout = 60000;
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行SingleSweep出错: " + Environment.NewLine + ex.Message);
            }
        }

        //频点单位MHz
        public ItemTestValues MeasMax(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasMax出错：" + Environment.NewLine + ex.Message);
            }


        }

        //频点单位MHz
        public ItemTestValues MeasMin(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasMin出错：" + Environment.NewLine + ex.Message);
            }
        }

        //频点单位MHz
        public ItemTestValues MeasRipple(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues MaxVal = MeasMax(itemTestRequir, itemTestInfor, testData);
            ItemTestValues MinVal = MeasMin(itemTestRequir, itemTestInfor, testData);

            ItemTestValues RippleVal;
            RippleVal.MeasuredValue = System.Math.Round(MaxVal.MeasuredValue - MinVal.MeasuredValue, 3);
            RippleVal.MeasuredFreq = 0;
            return RippleVal;
        }

        //频点单位MHz
        public ItemTestValues MeasVSWR(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                if (itemTestRequir.StarF == itemTestRequir.StopF)
                {
                    //设置Marker点
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:X " + itemTestRequir.StarF * 1000000);
                    //读取值
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                    XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                    XYValues.MeasuredFreq = System.Math.Round(itemTestRequir.StarF, 3);
                }
                else
                {
                    //设置Search Range
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                    //设置Marker Search
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                    //读取值

                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                    XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                    //读取频点
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                    XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);
                }
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行Meas驻波出错：" + Environment.NewLine + ex.Message);
            }
        }

        //频点单位MHz
        public ItemTestValues MeasPoint(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Marker点
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:X " + itemTestRequir.StarF * 1000000);
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                XYValues.MeasuredFreq = System.Math.Round(itemTestRequir.StarF, 3);
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasPoint出错：" + Environment.NewLine + ex.Message);
            }
        }

        //单位dBm
        public ItemTestValues Meas1dBIn(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            double MarkerTarget;
            double Data_Min;

            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = this.NaConnect.readNum();

                MarkerTarget = XYValues.MeasuredValue - 1;

                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");

                Data_Min = this.NaConnect.readNum();

                if (MarkerTarget > Data_Min)
                {
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TARG " + MarkerTarget);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.TARG);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                    XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                    XYValues.MeasuredFreq = 0;
                    MonitorSystem();
                    return XYValues;
                }
                else
                {
                    throw new Exception("Wrong 1dB measurment!");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("执行1dBIn出错：" + Environment.NewLine + ex.Message);
            }
        }

        //单位dBm
        public ItemTestValues Meas1dBOut(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            double MarkerTarget;
            double Data_Min;

            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = this.NaConnect.readNum();

                MarkerTarget = XYValues.MeasuredValue - 1;

                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");

                Data_Min = this.NaConnect.readNum();

                if (MarkerTarget > Data_Min)
                {
                    SelectTrace(itemTestRequir.Ch, itemTestRequir.Tr);
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TARG " + MarkerTarget);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.TARG);
                    this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                    XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                    this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");

                    XYValues.MeasuredValue += System.Math.Round(this.NaConnect.readNum(), 3);
                    XYValues.MeasuredFreq = 0;
                    MonitorSystem();
                    return XYValues;
                }
                else
                {
                    throw new Exception("Wrong 1dB measurment!");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("执行1dBOut出错：" + Environment.NewLine + ex.Message);
            }
        }

        [Obsolete("最初想出的方法，对itemTestRequir参数中的Interval有要求", true)]
        public ItemTestValues MeasAverageObsolete(ItemTestRequir itemTestRequir)
        {
            try
            {
                ItemTestValues XYValues;
                //状态文件起始频率
                double StateStartPoint;
                double StateStopPoint;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                string StrFreqL;
                string StrFreqH;

                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值
                double[,] MagPointsArray = null;

                SelectTrace(itemTestRequir.Ch, itemTestRequir.Tr);

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:STAR?");
                StateStartPoint = this.NaConnect.readNum() / 1000000;
                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:STOP?");
                StateStopPoint = this.NaConnect.readNum() / 1000000;

                // * Span的目的是取整，方便找到测试带宽起始频率在数组中的位置
                StrFreqL = Convert.ToString((itemTestRequir.StarF - StateStartPoint) * itemTestRequir.Interval);
                StrFreqH = Convert.ToString((itemTestRequir.StopF - StateStartPoint) * itemTestRequir.Interval);

                //如抛出异常说明Span指标设置不合理！
                FreqPosL = Int32.Parse(StrFreqL);
                FreqPosH = Int32.Parse(StrFreqH);

                //此处要求测试带宽必须为整数
                Int32 Gap = FreqPosH - FreqPosL;

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");

                NAFreqData = this.NaConnect.readStr().Split(',');
                //将读出的字符串数组填到FrequenciesList中
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');

                DataArray = new double[2, NAValueData.Length / 2];
                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                MagPointsArray = new double[Gap + 1, 2];
                Int32 Findex = 0;
                for (Int32 position = Convert.ToInt32(FreqPosL); position <= Convert.ToInt32(FreqPosH); position++)
                {
                    MagPointsArray[Findex, 0] = Frequencies[position];
                    MagPointsArray[Findex, 1] = DataArray[0, position];
                    Findex += 1;
                }
                double SumData = 0;
                double SumFreq = 0;
                for (Int32 index = 0; index <= Gap; index++)
                {
                    SumData = SumData + MagPointsArray[index, 1];
                    SumFreq = SumFreq + MagPointsArray[index, 0];
                }

                XYValues.MeasuredValue = Math.Round(SumData / (Gap + 1), 3);
                XYValues.MeasuredFreq = Math.Round(SumFreq / (Gap + 1), 3);
                return XYValues;
            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasAverage出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasAverage出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasAverage出错：" + Environment.NewLine + ex.Message);
            }

        }

        public ItemTestValues MeasAverage(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值
                List<mkPiont> ValidValue = new List<mkPiont>();

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");
                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");
                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }

                double AvgData = ValidValue.Average(item => item.yPos);
                double AvgFreq = ValidValue.Average(item => item.xPos);

                XYValues.MeasuredValue = Math.Round(AvgData, 3);
                XYValues.MeasuredFreq = Math.Round(AvgFreq / 1000000, 3);
                MonitorSystem();
                return XYValues;
            }


            catch (Exception ex)
            {
                throw new Exception("执行MeasAverage出错：" + Environment.NewLine + ex.Message);
            }

        }

        public ItemTestValues MeasDelay(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                double Delay = 0;
                double temp = 0;
                string interval = itemTestRequir.Interval.ToString();
                Int32 Gap = Int32.Parse(interval);

                List<mkPiont> ValidValue = new List<mkPiont>();


                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }
                if (Gap < ValidValue.Count())
                {
                    List<mkPiont> DivideValidVal = new List<mkPiont>();

                    //Interval 代表 测试带宽内，子带宽所占的Point数
                    for (Int32 divide = 0; divide < (ValidValue.Count() - itemTestRequir.Interval); divide += 1)
                    {
                        DivideValidVal = ValidValue.GetRange(divide, Convert.ToInt32(itemTestRequir.Interval) + 1);
                        temp = DivideValidVal.Max(item => item.yPos) - DivideValidVal.Min(item => item.yPos);
                        if (temp > Delay) Delay = temp;
                    }
                }
                else
                {
                    Delay = ValidValue.Max(item => item.yPos) - ValidValue.Min(item => item.yPos);
                }


                XYValues.MeasuredValue = Math.Round(Delay * 1000000000, 3);
                XYValues.MeasuredFreq = 0;
                MonitorSystem();
                return XYValues;
            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasDelay出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasDelay出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行MeasDelay出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasDelayPoint(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {


            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Marker点
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:X " + itemTestRequir.StarF * 1000000);
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum() * 1000000000, 3);
                XYValues.MeasuredFreq = System.Math.Round(itemTestRequir.StarF, 3);
                MonitorSystem();

                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasDelayPoint出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasTemp(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                InputManual inputTemp = new InputManual("请输入温度", itemTestRequir.StarF, itemTestRequir.StopF);
                inputTemp.ShowDialog();
                XYValues.MeasuredValue = System.Math.Round(Convert.ToDouble(inputTemp.TempData.Temperature), 3);
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }
            catch (NullReferenceException nRE)
            {
                throw new Exception("输入温度值取消: " + nRE.Message);
            }
        }

        public ItemTestValues MeasVoltage(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                InputManual inputTemp = new InputManual("请输入电压", itemTestRequir.StarF, itemTestRequir.StopF);
                inputTemp.ShowDialog();
                XYValues.MeasuredValue = System.Math.Round(Convert.ToDouble(inputTemp.TempData.Temperature), 3);
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }
            catch (NullReferenceException nRE)
            {
                throw new Exception("输入电压值取消: " + nRE.Message);
            }
        }

        public ItemTestValues MeasManualValue(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                InputManual inputTemp = new InputManual("请手动输入测量结果\n" + itemTestInfor.ItemName, itemTestInfor.LowSpec, itemTestInfor.HighSpec);
                inputTemp.ShowDialog();
                XYValues.MeasuredValue = System.Math.Round(Convert.ToDouble(inputTemp.TempData.Temperature), 3);
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }
            catch (NullReferenceException nRE)
            {
                throw new Exception("输入测量结果取消: " + nRE.Message);
            }
        }

        public ItemTestValues MeasToJPG(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            if (testData == null)
            {
                return new ItemTestValues { MeasuredFreq = 0, MeasuredValue = 0 };
            }
            try
            {
                Image printimg = this.DumpImage();

                while (!Directory.Exists(ActualPath.OutputDataPath))
                {
                    MessageBox.Show(ActualPath.OutputDataPath + "\n路径不存在，截图无法保存\n请检查！", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                string formatedSN = testData.Product.SerialNumber.Replace(":", "-").Replace(">", "").Replace("/", "").Replace(@"\", "").Replace(@"[", "").Replace(@" ", "").Replace(@"(", "").Replace(@")", "").Replace(@".", "");

                if (!Directory.Exists(ActualPath.OutputDataPath + @"\" + formatedSN))
                {
                    Directory.CreateDirectory(ActualPath.OutputDataPath + @"\" + formatedSN);
                }
                string dfName = itemTestInfor.ItemName;
                string msname = dfName.Substring(dfName.IndexOf(':') + 1, dfName.IndexOf('[') - dfName.IndexOf(':') - 1).Trim();

                string filename = ActualPath.OutputDataPath + @"\"
                    + formatedSN + @"\"
                    + itemTestInfor.IndexKey + "_"
                    + msname
                    + ".jpg";

                printimg.Save(filename, ImageFormat.Jpeg);
                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }

            catch (Exception ex)
            {
                throw new Exception("截图失败: " + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToS2P(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            if (testData == null)
            {
                return new ItemTestValues { MeasuredFreq = 0, MeasuredValue = 0 };
            }
            try
            {
                MemoryStream ms = this.SaveS2P();

                while (!Directory.Exists(ActualPath.OutputDataPath))
                {
                    MessageBox.Show(ActualPath.OutputDataPath + "\n路径不存在，S2P文件无法保存\n请检查！", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                string formatedSN = testData.Product.SerialNumber.Replace(":", "-").Replace(">", "").Replace("/", "").Replace(@"\", "").Replace(@"[", "").Replace(@" ", "").Replace(@"(", "").Replace(@")", "").Replace(@".", "");

                if (!Directory.Exists(ActualPath.OutputDataPath + @"\" + formatedSN))
                {
                    Directory.CreateDirectory(ActualPath.OutputDataPath + @"\" + formatedSN);
                }
                string dfName = itemTestInfor.ItemName;
                string msname = dfName.Substring(dfName.IndexOf(':') + 1, dfName.IndexOf('[') - dfName.IndexOf(':') - 1).Trim();

                string filename = ActualPath.OutputDataPath + @"\"
                    + formatedSN + @"\"
                    + itemTestInfor.IndexKey + "_"
                    + msname
                    + ".s2p";


                CommonAuxiliary.streamTofile(ms, filename);

                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("S2P文件保存失败: " + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiXMLS2P(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            if (testData == null)
            {
                return new ItemTestValues { MeasuredFreq = 0, MeasuredValue = 0 };
            }
            try
            {
                List<string> S2PDscription = new List<string>();
                MemoryStream ms = this.SaveS2P();

                using (StreamReader sr = new StreamReader(ms))
                {
                    while (sr.Peek() > -1)
                    {
                        S2PDscription.Add(sr.ReadLine());
                    }
                }
                HuaWeiDataStruct XMLS2P = new HuaWeiDataStruct(testData, S2PDscription, itemTestInfor);

                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("HuaWeiXMLS2P文件保存失败: " + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiJPG(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            if (testData == null)
            {
                return new ItemTestValues { MeasuredFreq = 0, MeasuredValue = 0 };
            }
            try
            {
                Image printimg = this.DumpImage();
                while (!Directory.Exists(ActualPath.OutputDataPath))
                {
                    MessageBox.Show(ActualPath.OutputDataPath + "\n路径不存在，截图无法保存\n请检查！", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                string savetime = DateTime.Now.ToString("s").Replace(":", "").Replace("-", "");
                string filename = ActualPath.OutputDataPath + @"\"
                    + testData.Product.SerialNumber + "_"
                    + testData.TestProgram.ProcessStep + "_"
                    + CommonAuxiliary.RePlaceSymbol(itemTestInfor.ItemName) + "_"
                    + savetime
                    + ".jpg";
                printimg.Save(filename, ImageFormat.Jpeg);

                while (!Directory.Exists(HuaWeiDataStruct.HuaWeiDataPath))
                {
                    MessageBox.Show(HuaWeiDataStruct.HuaWeiDataPath + "\n路径不存在，截图无法上传给HuaWei\n请检查！", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                filename = HuaWeiDataStruct.HuaWeiDataPath
                          + testData.Product.CustomerSN + "_"
                          + savetime
                          + ".jpg";
                printimg.Save(filename, ImageFormat.Jpeg);

                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }

            catch (Exception ex)
            {
                throw new Exception("HuaWei截图失败: " + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiS2P(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            if (testData == null)
            {
                return new ItemTestValues { MeasuredFreq = 0, MeasuredValue = 0 };
            }
            try
            {
                MemoryStream ms = this.SaveS2P();

                while (!Directory.Exists(ActualPath.OutputDataPath))
                {
                    MessageBox.Show(ActualPath.OutputDataPath + "\n路径不存在，S2P文件无法保存\n请检查！", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                string savetime = DateTime.Now.ToString("s").Replace(":", "").Replace("-", "");
                string filename = ActualPath.OutputDataPath + @"\"
                    + testData.Product.SerialNumber + "_"
                    + testData.TestProgram.ProcessStep + "_"
                    + CommonAuxiliary.RePlaceSymbol(itemTestInfor.ItemName) + "_"
                    + savetime
                    + ".s2p";
                CommonAuxiliary.streamTofile(ms, filename);

                while (!Directory.Exists(HuaWeiDataStruct.HuaWeiDataPath))
                {
                    MessageBox.Show(HuaWeiDataStruct.HuaWeiDataPath + "\n路径不存在，S2P文件无法上传给HuaWei\n请检查！", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                filename = HuaWeiDataStruct.HuaWeiDataPath
                          + testData.Product.CustomerSN + "_"
                          + savetime
                          + ".s2p";
                CommonAuxiliary.streamTofile(ms, filename);

                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("HuaWeiS2P文件保存失败: " + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiXMLIL(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                List<mkPiont> ValidValue = new List<mkPiont>();

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');

                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position] / 1000000, yPos = DataArray[0, position] });
                }

                Channel channel = new Channel();
                channel.ChannelName = itemTestInfor.ItemName.Substring(itemTestInfor.ItemName.IndexOf(':') + 1, itemTestInfor.ItemName.IndexOf('[') - itemTestInfor.ItemName.IndexOf(':') - 1).Trim();
                channel.FreqStep = itemTestRequir.Interval;
                channel.StartToStopFreq = itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0");
                List<ValueSpec> ILValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();
                for (double freq = itemTestRequir.StarF; Math.Round(freq, 1) <= itemTestRequir.StopF; freq += itemTestRequir.Interval)
                {
                    tempVal.MeasuredFreq = Math.Round(freq, 1);

                    //tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                    tempVal.Spc = Math.Abs(itemTestInfor.HighSpec).ToString("0.00") + ":" + Math.Abs(itemTestInfor.LowSpec).ToString("0.00");
                    tempVal.MeasuredValue = Math.Round(SearchCharData(tempVal.MeasuredFreq, ValidValue), 2);
                    tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                    ILValues.Add(tempVal);
                }
                channel.ChannelStatus = ResultStatus.Passed;
                foreach (ValueSpec val in ILValues)
                {
                    if (val.result == ResultStatus.Failed)
                    {
                        channel.ChannelStatus = ResultStatus.Failed;
                        break;
                    }
                }
                channel.Values = ILValues;

                if (testData != null)
                {
                    TempDatas.RecordData("Insertion Loss Test", channel);
                }

                XYValues.MeasuredValue = ILValues.Min(val => val.MeasuredValue);
                XYValues.MeasuredFreq = ILValues.First(val => val.MeasuredValue == XYValues.MeasuredValue).MeasuredFreq;
                MonitorSystem();
                return XYValues;

            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasToHuaWeiXMLIL出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasToHuaWeiXMLIL出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行MeasToHuaWeiXMLIL出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiXMLDelay(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                List<mkPiont> ValidValue = new List<mkPiont>();


                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');

                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position] / 1000000, yPos = DataArray[0, position] });
                }

                Channel channel = new Channel();
                channel.ChannelName = itemTestInfor.ItemName.Substring(itemTestInfor.ItemName.IndexOf(':') + 1, itemTestInfor.ItemName.IndexOf('[') - itemTestInfor.ItemName.IndexOf(':') - 1).Trim();
                channel.FreqStep = itemTestRequir.Interval;
                channel.StartToStopFreq = itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0");
                List<ValueSpec> DelayValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();
                for (double freq = itemTestRequir.StarF; Math.Round(freq, 1) <= itemTestRequir.StopF; freq += itemTestRequir.Interval)
                {
                    tempVal.MeasuredFreq = Math.Round(freq, 1);
                    tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                    tempVal.MeasuredValue = Math.Round(SearchCharData(tempVal.MeasuredFreq, ValidValue) * 1000000000);
                    tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                    DelayValues.Add(tempVal);
                }
                channel.ChannelStatus = ResultStatus.Passed;
                foreach (ValueSpec val in DelayValues)
                {
                    if (val.result == ResultStatus.Failed)
                    {
                        channel.ChannelStatus = ResultStatus.Failed;
                        break;
                    }
                }
                channel.Values = DelayValues;

                if (testData != null)
                {
                    TempDatas.RecordData("Group Delay Test", channel);
                }

                XYValues.MeasuredValue = DelayValues.Max(val => val.MeasuredValue);
                XYValues.MeasuredFreq = DelayValues.First(val => val.MeasuredValue == XYValues.MeasuredValue).MeasuredFreq;
                MonitorSystem();
                return XYValues;
            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasToHuaWeiXMLDelay出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasToHuaWeiXMLDelay出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行MeasToHuaWeiXMLDelay出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiExcelDelay(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                List<mkPiont> ValidValue = new List<mkPiont>();


                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');

                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position] / 1000000, yPos = DataArray[0, position] });
                }

                Channel channel = new Channel();
                channel.ChannelName = itemTestInfor.ItemName.Substring(itemTestInfor.ItemName.IndexOf(':') + 1, itemTestInfor.ItemName.IndexOf('[') - itemTestInfor.ItemName.IndexOf(':') - 1).Trim();
                channel.FreqStep = itemTestRequir.Interval;
                channel.StartToStopFreq = itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0");
                List<ValueSpec> DelayValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();
                int stepcount = 0;
                for (double freq = itemTestRequir.StarF; Math.Round(freq, 1) <= itemTestRequir.StopF; freq += itemTestRequir.Interval)
                {
                    stepcount++;
                    if (stepcount == 2 && itemTestRequir.StarF % 1 != 0 && itemTestRequir.Interval % 1 == 0)
                    {
                        tempVal.MeasuredFreq = Math.Ceiling(itemTestRequir.StarF);
                        freq = Math.Ceiling(itemTestRequir.StarF);
                    }
                    tempVal.MeasuredFreq = Math.Round(freq, 1);
                    tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                    tempVal.MeasuredValue = Math.Round(SearchCharData(tempVal.MeasuredFreq, ValidValue) * 1000000000, 2);
                    tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                    DelayValues.Add(tempVal);

                    if (tempVal.MeasuredFreq + itemTestRequir.Interval > itemTestRequir.StopF && tempVal.MeasuredFreq != itemTestRequir.StopF)
                    {
                        tempVal.MeasuredFreq = itemTestRequir.StopF;
                        tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                        tempVal.MeasuredValue = Math.Round(SearchCharData(tempVal.MeasuredFreq, ValidValue) * 1000000000, 2);
                        tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                        DelayValues.Add(tempVal);
                    }
                }
                channel.ChannelStatus = ResultStatus.Passed;
                foreach (ValueSpec val in DelayValues)
                {
                    if (val.result == ResultStatus.Failed)
                    {
                        channel.ChannelStatus = ResultStatus.Failed;
                        break;
                    }
                }
                channel.Values = DelayValues;

                if (testData != null)
                {
                    TempExcelDatas.RecordData("Group Delay", channel);
                }

                XYValues.MeasuredValue = DelayValues.Max(val => val.MeasuredValue);
                XYValues.MeasuredFreq = DelayValues.First(val => val.MeasuredValue == XYValues.MeasuredValue).MeasuredFreq;
                MonitorSystem();
                return XYValues;
            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasToHuaWeiExcelDelay出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasToHuaWeiExcelDelay出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行MeasToHuaWeiExcelDelay出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasAluAvgReg(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                List<mkPiont> ValidValue = new List<mkPiont>();


                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');

                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }

                List<ValueSpec> PointValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();
                for (double freq = itemTestRequir.StarF; freq <= itemTestRequir.StopF; freq += itemTestRequir.Interval)
                {
                    tempVal.MeasuredFreq = freq;
                    tempVal.MeasuredValue = SearchCharData(tempVal.MeasuredFreq * 1000000, ValidValue);
                    PointValues.Add(tempVal);
                }

                List<double> powValues = new List<double>();
                double val1 = 0;
                double val2 = 0;
                double powval = 0;
                for (int i = 0; i < PointValues.Count - 1; i++)
                {
                    val1 = PointValues[i].MeasuredValue;
                    val2 = PointValues[i + 1].MeasuredValue;
                    powval = 2 / (1 / Math.Pow(10, val1 / 10) + 1 / Math.Pow(10, val2 / 10));
                    powValues.Add(powval);
                }

                XYValues.MeasuredValue = Math.Log10(powValues.Sum() / powValues.Count()) * 10;
                XYValues.MeasuredFreq = 0;
                MonitorSystem();
                return XYValues;

            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasAluAvgReg出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasAluAvgReg出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行MeasAluAvgReg出错：" + Environment.NewLine + ex.Message);
            }
        }

        private double SearchCharData(double SearchFreq, List<mkPiont> ValidValue)
        {
            mkPiont pSample = new mkPiont();
            foreach (mkPiont cSample in ValidValue)
            {
                if (cSample.xPos > SearchFreq & pSample.xPos == 0)
                {
                    return -1;
                }
                else if (cSample.xPos == SearchFreq)
                {
                    return cSample.yPos;
                }
                else if (Convert.ToDouble(cSample.xPos) > SearchFreq)
                {
                    double cFreq = cSample.xPos;
                    double pFreq = pSample.xPos;
                    double ratio = (cFreq - SearchFreq) / (cFreq - pFreq);
                    double cVal = cSample.yPos;
                    double pVal = pSample.yPos;

                    return cVal + ((pVal - cVal) * ratio);
                }

                pSample = cSample;
            }
            return -1;
        }

        public ItemTestValues MeasILRipple(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                double Ripple = 0;
                double temp = 0;
                string interval = itemTestRequir.Interval.ToString();
                Int32 Gap = Int32.Parse(interval);

                List<mkPiont> ValidValue = new List<mkPiont>();


                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }
                if (Gap < ValidValue.Count())
                {
                    List<mkPiont> DivideValidVal = new List<mkPiont>();

                    //Interval 代表 测试带宽内，子带宽所占的Point数
                    for (Int32 divide = 0; divide < (ValidValue.Count() - itemTestRequir.Interval); divide += 1)
                    {
                        DivideValidVal = ValidValue.GetRange(divide, Convert.ToInt32(itemTestRequir.Interval) + 1);
                        temp = DivideValidVal.Max(item => item.yPos) - DivideValidVal.Min(item => item.yPos);
                        if (temp > Ripple) Ripple = temp;
                    }
                }
                else
                {
                    Ripple = ValidValue.Max(item => item.yPos) - ValidValue.Min(item => item.yPos);
                }


                XYValues.MeasuredValue = Math.Round(Ripple, 3);
                XYValues.MeasuredFreq = 0;
                MonitorSystem();
                return XYValues;
            }
            catch (FormatException FEx)
            {
                throw new Exception("执行ILRipple出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行ILRipple出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行ILRipple出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasActiveRej(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                if (testData == null)
                {
                    throw new Exception("执行MeasActiveRej出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }

                if (!testData.CheckTestIndexKey(Convert.ToInt32(itemTestRequir.Interval)))
                {
                    throw new Exception("执行MeasActiveRej出错：" + Environment.NewLine + "第" + itemTestRequir.Interval.ToString() + "步可能还未测试,无法调用它的测试信息!");
                }

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");

                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum() - testData.GetTestValue(Convert.ToInt32(itemTestRequir.Interval)), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);


                MonitorSystem();

                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasActiveRej出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasDelayMax(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum() * 1000000000, 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasDelayMax出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasBalance(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                string[] IndexSteps = itemTestRequir.State.Split(',');
                List<double> testValues = new List<double>();
                if (testData == null)
                {
                    throw new Exception("执行MeasBalance出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }
                foreach (string strindex in IndexSteps)
                {
                    if (!testData.CheckTestIndexKey(Convert.ToInt32(strindex)))
                    {
                        throw new Exception("执行MeasBalance出错：" + Environment.NewLine + "第" + strindex + "步可能还未测试,无法调用它的测试信息!");
                    }

                    testValues.Add(testData.GetTestValue(Convert.ToInt32(strindex)));
                }

                XYValues.MeasuredValue = System.Math.Round(testValues.Max() - testValues.Min(), 3);
                XYValues.MeasuredFreq = 0;

                MonitorSystem();

                return XYValues;

            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasBalance出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasPointActiveAttenuation(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                if (testData == null)
                {
                    throw new Exception("执行MeasPointActiveAttenuation出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }

                if (!testData.CheckTestIndexKey(Convert.ToInt32(itemTestRequir.Interval)))
                {
                    throw new Exception("执行MeasPointActiveAttenuation出错：" + Environment.NewLine + "因为第" + itemTestRequir.Interval.ToString() + "步未测试,无法调用它的测试信息!");
                }

                //执行除法
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC DIV");
                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Marker点
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:X " + itemTestRequir.StarF * 1000000);
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                XYValues.MeasuredFreq = System.Math.Round(itemTestRequir.StarF, 3);

                //还原
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC NORM");

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasPointActiveAttenuation出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasActiveAttenuationMax(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                if (testData == null)
                {
                    throw new Exception("执行MeasActiveAttenuationMax出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }

                if (!testData.CheckTestIndexKey(Convert.ToInt32(itemTestRequir.Interval)))
                {
                    throw new Exception("执行MeasActiveAttenuationMax出错：" + Environment.NewLine + "因为第" + itemTestRequir.Interval.ToString() + "步未测试,无法调用它的测试信息!");
                }


                //执行除法
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC DIV");

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");

                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);

                //还原
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC NORM");

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasActiveAttenuationMax出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasActiveAttenuationMin(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                if (testData == null)
                {
                    throw new Exception("执行MeasActiveAttenuationMin出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }

                if (!testData.CheckTestIndexKey(Convert.ToInt32(itemTestRequir.Interval)))
                {
                    throw new Exception("执行MeasActiveAttenuationMin出错：" + Environment.NewLine + "因为第" + itemTestRequir.Interval.ToString() + "步未测试,无法调用它的测试信息!");
                }

                //执行除法
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC DIV");

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");

                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);

                //还原
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC NORM");

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasActiveAttenuationMin出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasMaxAndDataMemory(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {

            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);

                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MATH:MEM");
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasMaxAndDataMemory出错：" + Environment.NewLine + ex.Message);
            }

        }

        public ItemTestValues MeasMinAndDataMemory(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);

                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MATH:MEM");
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasMinAndDataMemory出错：" + Environment.NewLine + ex.Message);
            }
        }

        # endregion

        # region "SOLT 校准"



        public List<string> CalIni(Int32 chNum, CalibrationInfor calibrationInfor)
        {
            try
            {
                List<string> StepInfor = new List<string>();
                ClearError();
                SelectChannel(chNum);
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":CORR:COLL:CKIT " + CalKitLookUp[calibrationInfor.CalKit]);
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":CORR:COLL:METH:" + CalMethodLookUp[calibrationInfor.CalMethod] + "2 1,2");

                switch (calibrationInfor.CalMethod)
                {
                    case "Solt2PortCal":
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port1");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port1");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port1");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port2");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port2");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port2");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port1 and Port2");

                        break;
                    case "TRL2PortCal":
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " ReflectComponent to Port1");
                        StepInfor.Add("Connect " + calibrationInfor.CalKit + " ReflectComponent to Port2");
                        StepInfor.Add("Connect Port1 to Port2");
                        StepInfor.Add("Connect INSERTABLE THRU STANDARD between Port1 and Port2");

                        break;

                    default:
                        throw new Exception(calibrationInfor.CalMethod + " 该校准方式尚未定义，请检查指标配置!");

                }
                MonitorSystem();
                return StepInfor;

            }
            catch (Exception ex)
            {
                throw new Exception("设置校准件信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        public List<string> CalIni(List<ChanInfor> CiList)
        {
            try
            {
                List<string> StepInfor = new List<string>();
                ClearError();
                string _calkit = CiList[0].Calibration.CalKit;
                string _calmethod = CiList[0].Calibration.CalMethod;

                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    SelectChannel(chaninfor.Chan);
                    this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:CKIT " + CalKitLookUp[_calkit]);
                    this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:METH:" + CalMethodLookUp[_calmethod] + "2 1,2");
                });

                switch (CiList[0].Calibration.CalMethod)
                {
                    case "Solt2PortCal":
                        StepInfor.Add("Connect " + _calkit + " OPEN to Port1");
                        StepInfor.Add("Connect " + _calkit + " SHORT to Port1");
                        StepInfor.Add("Connect " + _calkit + " LOAD to Port1");
                        StepInfor.Add("Connect " + _calkit + " OPEN to Port2");
                        StepInfor.Add("Connect " + _calkit + " SHORT to Port2");
                        StepInfor.Add("Connect " + _calkit + " LOAD to Port2");
                        StepInfor.Add("Connect " + _calkit + " TO " + _calkit + " ADAPTER between Port1 and Port2");

                        break;
                    case "TRL2PortCal":
                        StepInfor.Add("Connect " + _calkit + " ReflectComponent to Port1");
                        StepInfor.Add("Connect " + _calkit + " ReflectComponent to Port2");
                        StepInfor.Add("Connect Port1 to Port2");
                        StepInfor.Add("Connect INSERTABLE THRU STANDARD between Port1 and Port2");

                        break;

                    default:
                        throw new Exception(_calmethod + " 该校准方式尚未定义，请检查指标配置!");

                }
                MonitorSystem();
                return StepInfor;

            }
            catch (Exception ex)
            {
                throw new Exception("设置校准件信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step1(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:OPEN 1");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Port1 Open 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step2(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SHOR 1");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();

            }
            catch (Exception ex)
            {
                throw new Exception("执行Port1 Short 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step3(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                //SUBC使得网分有记忆功能，而Load校准不需要SUB信息，所以让它默认
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC 1");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:LOAD 1");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();


            }
            catch (Exception ex)
            {
                throw new Exception("执行Port1 Load 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step4(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:OPEN 2");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Port2 Open 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step5(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SHOR 2");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Port2 Short 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step6(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC 1");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:LOAD 2");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();

            }
            catch (Exception ex)
            {
                throw new Exception("执行Port2 Load 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Solt2Port_Step7(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 480000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC 1");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:THRU 1,2");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:THRU 2,1");
                    this.NaConnect.timeout = 60000;
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Thru 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CalFinish(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC 1");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SAVE");
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行校准Done出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CalAbort(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:CLE");
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("退出校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        # endregion

        # region "SOLT 校准"

        public void Trl2Port_Step1(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:TRLR 1");
                    this.NaConnect.timeout = 60000;
                });

                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Port1 Reflect 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Trl2Port_Step2(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 300000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:TRLR 2");
                    this.NaConnect.timeout = 60000;
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Port2 Reflect 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Trl2Port_Step3(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
               {
                   this.NaConnect.timeout = 480000;
                   this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC 1");
                   this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:TRLT 1,2");
                   this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:TRLT 2,1");
                   this.NaConnect.timeout = 60000;
               });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Thru/Line 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void Trl2Port_Step4(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor ch)
                {
                    this.NaConnect.timeout = 480000;
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:SUBC 1");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:TRLL 1,2");
                    this.ReWriteNaConnect.Write(@":SENS" + ch.Chan + ":CORR:COLL:TRLL 2,1");
                    this.NaConnect.timeout = 60000;
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Line/Match 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        # endregion


        //返回是否网分连接状态
        //重写命令绑定有效
        public bool open()
        {
            if (this.NaConnect.open())
            {
                this.NaConnect.timeout = 60000;
                ReWriteNaConnect = new ReWriteCommand(NaConnect);
                if (this.AvailableENA(this.NetWorkAnalyzerSerialNumber))
                {
                    return true;
                }
                else
                {
                    throw new Exception("不能识别此网分序列号!\n" + this.NetWorkAnalyzerSerialNumber);
                }
            }
            else
            {
                return false;
            }
        }

        //关闭网分通信连接
        public void close()
        {
            this.NaConnect.Close();
        }

        public void CallState(string pth, string filename, string transferpth)
        {
            try
            {
                ClearError();
                if (CatalogExistFile(pth, filename))
                {
                    this.NaConnect.sendCmd(@":MMEM:DEL '" + pth + @"\" + filename + ".sta'");
                }
                byte[] byteData = CommonAuxiliary.filetobytes(transferpth + @"\" + this.NetWorkAnalyzerSerialNumber + "_" + filename + ".sta");
                string sHeader = ":MMEM:TRAN '" + pth + @"\" + filename + ".sta" + "',";
                this.NaConnect.sendBlock(sHeader, byteData);
                this.ReWriteNaConnect.Write(@":MMEM:LOAD '" + pth + @"\" + filename + ".sta'");
                this.NaConnect.sendCmd(@":MMEM:DEL '" + pth + @"\" + filename + ".sta'");

                int getchCount = 0;
                this.NaConnect.sendCmd(@":DISP:SPL?");
                string disp = this.NaConnect.readStr().Trim();
                List<int> num = (from ChanNum in ChanLookUp
                                 where ChanNum.Value == disp
                                 select ChanNum.Key).ToList();
                getchCount = num[0];

                for (int ch = 1; ch <= getchCount; ch++)
                {
                    this.ReWriteNaConnect.Write(@"DISP:MAX ON");
                    this.ReWriteNaConnect.Write(@"DISP:WIND" + ch + ":MAX ON");
                }


                MonitorSystem();
            }

            catch (FileNotFoundException FNe)
            {
                throw new Exception(FNe.Message + "\n未发现指定路径下状态文件!\n请重新设置校验!");
            }

            catch (Exception ex)
            {
                throw new Exception("调用网分状态文件出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void SaveState(string pth, string filename, string transferpth)
        {
            try
            {
                ClearError();
                this.MakeDirectory(pth);
                this.ReWriteNaConnect.Write(@":MMEM:STOR:STYP CST");
                this.ReWriteNaConnect.Write(@":MMEM:STOR '" + pth + @"\" + filename + ".sta'");
                this.NaConnect.sendCmd(":MMEM:TRAN? " + "'" + pth + @"\" + filename + ".sta" + "'");
                MemoryStream ms = new MemoryStream((byte[])this.NaConnect.readByte());
                CommonAuxiliary.streamTofile(ms, transferpth + @"\" + this.NetWorkAnalyzerSerialNumber + "_" + filename + ".sta");
                this.NaConnect.sendCmd(@":MMEM:DEL '" + pth + @"\" + filename + ".sta'");

                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("保存网分状态文件出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void MakeDirectory(string pth)
        {
            try
            {
                List<string> Dirs = new List<string>();
                Dirs = CommonAuxiliary.DivideDirectory(pth);
                for (int dircount = 0; dircount < Dirs.Count; dircount++)
                {
                    if (dircount == 0)
                    {
                        if (CatalogExistFile(Dirs[dircount].Substring(0, Dirs[dircount].LastIndexOf("\\") + 1), Dirs[dircount].Substring(Dirs[dircount].LastIndexOf("\\") + 1) + "\\") == false)
                        {
                            this.ReWriteNaConnect.Write(":MMEM:MDIR " + "\"" + Dirs[dircount] + "\"");
                        }
                    }
                    else
                    {
                        if (CatalogExistFile(Dirs[dircount - 1], Dirs[dircount].Substring(Dirs[dircount].LastIndexOf("\\") + 1) + "\\") == false)
                        {
                            this.ReWriteNaConnect.Write(":MMEM:MDIR " + "\"" + Dirs[dircount] + "\"");
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                throw new Exception("创建网分状态文件路径出错: " + Environment.NewLine + ex.Message);
            }
        }


        public List<string> CatalogStaFile(string pth)
        {
            try
            {
                List<string> StaFiles = new List<string>();
                this.NaConnect.sendCmd(@"MMEM:CAT? '" + pth + "'");
                StaFiles = this.NaConnect.readStr().Split(new char[] { ',', ',' }).ToList<string>();
                StaFiles.RemoveAll(n => n.Contains(".STA") == false);
                return StaFiles;
            }
            catch (Exception ex)
            {
                throw new Exception("查询网分目录失败: " + Environment.NewLine + ex.Message);
            }
        }

        public string NetWorkAnalyzerName
        {
            get
            {
                return this.NaName;
            }
        }

        public bool Connected
        {
            get
            {
                return this.NaConnect.IsConnected;
            }
        }

        public string NetWorkAnalyzerSerialNumber
        {
            get
            {
                try
                {
                    this.NaConnect.sendCmd(@"*IDN?");
                    string Sn = this.NaConnect.readStr().Split(',')[2];
                    return Sn;
                }
                catch (Exception ex)
                {
                    throw new Exception("获取仪器序列号出错: " + Environment.NewLine + ex.Message);
                }
            }
        }

        private void HoldAllChannels(int chNum)
        {
            try
            {
                if (chNum < 1) throw new Exception("Channel数量不应该小于1");

                for (int count = 1; count <= chNum; count++)
                {
                    this.ReWriteNaConnect.Write(@":INIT" + count + ":CONT OFF");
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Hold所有通道出错: " + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// 切记：可能会造成死循环！
        /// </summary>
        /// <returns></returns>
        public SystemError GetSystemError()
        {
            try
            {
                SystemError systemerr = new SystemError() { Error = "", HaveError = false };
                string[] errArray;
                double errNum = 0;
                string errstr = "";
                StringBuilder errbuilder = new StringBuilder("");

                this.NaConnect.sendCmd(":SYST:ERR?");
                systemerr.Error = this.NaConnect.readStr();
                errArray = systemerr.Error.Split(',');
                errNum = Convert.ToDouble(errArray[0]);
                if (errNum == 0)
                {
                    return systemerr;
                }
                else
                {
                    systemerr.HaveError = true;
                    errbuilder.Append(errArray[1]);
                    while (errNum != 0)
                    {
                        this.NaConnect.sendCmd(":SYST:ERR?");
                        errstr = this.NaConnect.readStr();
                        errArray = errstr.Split(',');
                        errNum = Convert.ToDouble(errArray[0]);
                        if (errNum != 0)
                        {
                            errbuilder.Append(Environment.NewLine + errstr);
                        }
                        systemerr.Error = errbuilder.ToString();
                    }
                    return systemerr;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("获取系统错误信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void ClearError()
        {
            try
            {
                this.ReWriteNaConnect.Write("*CLS");
            }
            catch (Exception ex)
            {
                throw new Exception("执行清空寄存器错误信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        private void MonitorSystem()
        {
            SystemError sysErr = GetSystemError();
            if (sysErr.HaveError)
            {
                throw new Exception("执行命令出错: " + sysErr.Error);
            }
        }

        public void CheckMark(int chNum, MarkerSetInfor markSetInfor)
        {
            //throw new NotImplementedException();
        }

        private Boolean CatalogExistFile(string pth, string fileName)
        {
            try
            {
                string FileDir;
                this.NaConnect.sendCmd(@"MMEM:CAT? '" + pth + "'");
                FileDir = this.NaConnect.readStr().ToUpper();
                string[] filedirs = FileDir.Split(',');

                if (filedirs.Contains(fileName.ToUpper()))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("查询网分目录失败: " + Environment.NewLine + ex.Message);
            }
        }

        private Boolean AvailableENA(string Sn)
        {
            List<string> AvailEna = new List<string>();

            //AvailEna.Add("MY46631755");

            
            //if (AvailEna.Contains(Sn))
            //{
                return true;
            //}
            //else
            //{
            //    return false;
            //}
        }

        public string TRANCalIni(int chNum, CalibrationInfor calibrationInfor)
        {
            try
            {
                string StepInfor;
                ClearError();
                SelectChannel(chNum);

                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":CORR:COLL:CKIT " + CalKitLookUp[calibrationInfor.CalKit]);
                this.ReWriteNaConnect.Write(@":SENS" + chNum + ":CORR:COLL:METH:" + CalMethodLookUp[calibrationInfor.CalMethod] + " 1,2");

                StepInfor = "Connect " + calibrationInfor.CalKit + " THRU" + " between Port1 and Port2";

                MonitorSystem();
                return StepInfor;

            }
            catch (Exception ex)
            {
                throw new Exception("设置校准件信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        public string TRANCalIni(List<ChanInfor> CiList)
        {
            try
            {
                string StepInfor;
                ClearError();

                string _calkit = CiList[0].Calibration.CalKit;
                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    SelectChannel(chaninfor.Chan);
                });


                StepInfor = "Connect " + _calkit + " THRU" + " between Port1 and Port2";

                MonitorSystem();
                return StepInfor;

            }
            catch (Exception ex)
            {
                throw new Exception("设置校准件信息出错: " + Environment.NewLine + ex.Message);
            }
        }
        public void TRAN_Step1(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                ClearError();
                string _calkit = CiList[0].Calibration.CalKit;
                string _calmethod = CiList[0].Calibration.CalMethod;
                CiList.ForEach(delegate(ChanInfor chaninfor)
                    {
                        SelectChannel(chaninfor.Chan);
                        this.NaConnect.timeout = 480000;
                        foreach (TraceInfor ti in chaninfor.Traces)
                        {
                            if ((ti.Measurement != SPar.S21) && (ti.Measurement != SPar.S12))
                            {
                                throw new Exception("Measurement为 " + ti.Measurement + " 时不可以执行Thru校准");
                            }
                        }
                        if (chaninfor.Traces.Count < 1)
                        {
                            throw new Exception("状态文件指标Trace数量设置错误!");
                        }
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:CKIT " + CalKitLookUp[_calkit]);
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:METH:" + CalMethodLookUp[_calmethod] + " 1,2");
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 1,2");
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SAVE");

                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:CKIT " + CalKitLookUp[_calkit]);
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:METH:" + CalMethodLookUp[_calmethod] + " 2,1");
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 2,1");
                        this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SAVE");
                        this.NaConnect.timeout = 60000;
                    });

                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行Thru 校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        //与CalFinish一样
        public void TRAN_CalFinish(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            try
            {
                //不需要
            }
            catch (Exception ex)
            {
                throw new Exception("执行校准Done出错: " + Environment.NewLine + ex.Message);
            }
        }

        public Image DumpImage()
        {
            try
            {
                this.MakeDirectory(@"C:\SParTest");
                this.ReWriteNaConnect.Write("MMEM:STOR:IMAG " + "'" + @"c:/SParTest/myFile.png" + "'");
                return this.NaConnect.readImage(@"c:/SParTest/myFile.png");
            }
            catch (Exception ex)
            {
                throw new Exception("截图失败: " + Environment.NewLine + ex.Message);
            }
        }

        public MemoryStream SaveS2P()
        {
            try
            {
                this.ReWriteNaConnect.Write(":MMEM:STOR:SNP:TYPE:S2P 1,2");
                this.MakeDirectory(@"C:\SParTest");
                this.ReWriteNaConnect.Write(":MMEM:STOR:SNP " + "'" + @"C:\SPARTEST\mys2p.s2p" + "'");
                this.NaConnect.sendCmd(":MMEM:TRAN? " + "'" + @"c:/SParTest/mys2p.s2p" + "'");
                MemoryStream ms = new MemoryStream((byte[])this.NaConnect.readByte());
                return ms;
            }
            catch (Exception ex)
            {
                throw new Exception("截取S2P文件失败: " + Environment.NewLine + ex.Message);
            }
        }

        public void SetTraceSmooth(int chNum, int TrNum, SmoothInfor smoothInfor)
        {
            try
            {
                //之前有很多完成的项目文件没有此项设置，所以没有设置的都默认OFF
                if (smoothInfor == null)
                {
                    this.ReWriteNaConnect.Write(@":CALC" + chNum + ":SMO:STAT " + IsEnable.OFF);
                }
                else
                {
                    this.ReWriteNaConnect.Write(@":CALC" + chNum + ":SMO:STAT " + smoothInfor.SmoothingEnable);
                    this.ReWriteNaConnect.Write(@":CALC" + chNum + ":SMO:APER " + smoothInfor.SmoAperture);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("设置迹线Smooth出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void CheckTraceSmooth(int chNum, int TrNum, SmoothInfor smoothInfor)
        {
            try
            {
                int smoenable;
                double smoaper;
                if (smoothInfor == null)
                {
                    this.NaConnect.sendCmd(@":CALC" + chNum + ":SMO:STAT?");
                    smoenable = Convert.ToInt32(this.NaConnect.readNum());
                    if ((IsEnable)smoenable != IsEnable.OFF)
                    {
                        throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " SmoothingEnable错误: 默认为 OFF " + ";仪器读数为 " + (IsEnable)smoenable);
                    }
                }
                else
                {
                    this.NaConnect.sendCmd(@":CALC" + chNum + ":SMO:STAT?");
                    smoenable = Convert.ToInt32(this.NaConnect.readNum());
                    if ((IsEnable)smoenable != smoothInfor.SmoothingEnable)
                    {
                        throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " SmoothingEnable错误: 指标为 " + smoothInfor.SmoothingEnable + " ;仪器读数为 " + (IsEnable)smoenable);
                    }

                    this.NaConnect.sendCmd(@":CALC" + chNum + ":SMO:APER?");
                    smoaper = this.NaConnect.readNum();
                    if (smoaper != smoothInfor.SmoAperture)
                    {
                        throw new Exception("通道 " + chNum + " 迹线 " + TrNum + " SmoAperture错误: 指标为 " + smoothInfor.SmoAperture + " ;仪器读数为 " + smoaper);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Smooth信息异常: " + Environment.NewLine + ex.Message);
            }
        }

        public void Forbid_Peripheral(IsEnable value)
        {
            try
            {

                //if (value == IsEnable.ON)
                //{
                //    this.ReWriteNaConnect.Write(@":DISP:ENAB " + IsEnable.OFF);
                //}
                //else
                //{
                //    this.ReWriteNaConnect.Write(@":DISP:ENAB " + IsEnable.ON);
                //}
                this.ReWriteNaConnect.Write(@":SYST:KLOC:MOUS " + value);
                this.ReWriteNaConnect.Write(@":SYST:KLOC:KBD " + value);
            }

            catch (Exception ex)
            {
                throw new Exception("Set Peripheral Devices Local Status Fail :" + Environment.NewLine + ex.Message);
            }
        }

        public List<string> _3PortCalIni(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            try
            {
                List<string> StepInfor = new List<string>();
                ClearError();
                string _calkit = CiList[0].Calibration.CalKit;
                string _calmethod = CiList[0].Calibration.CalMethod;

                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    SelectChannel(chaninfor.Chan);
                    this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:CKIT " + CalKitLookUp[_calkit]);
                    this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:METH:" + CalMethodLookUp[_calmethod] + "3 1,2,3");
                });

                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port1");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port1");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port1");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port1 and Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port1 and Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port2 and Port3");
                MonitorSystem();
                return StepInfor;

            }
            catch (Exception ex)
            {
                throw new Exception("设置校准件信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void _3Port_Step(List<ChanInfor> CiList, Gender ConnectorGender, int Step)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    SelectChannel(chaninfor.Chan);
                    this.NaConnect.timeout = 300000;
                    switch (Step)
                    {
                        case 1:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 1");
                            break;
                        case 2:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 1");
                            break;
                        case 3:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 1");
                            break;
                        case 4:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 2");
                            break;
                        case 5:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 2");
                            break;
                        case 6:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 2");
                            break;
                        case 7:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 3");
                            break;
                        case 8:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 3");
                            break;
                        case 9:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 3");
                            break;
                        case 10:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 1,2");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 2,1");
                            break;
                        case 11:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 1,3");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 3,1");
                            break;
                        case 12:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 2,3");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 3,2");
                            break;
                    }
                    this.NaConnect.timeout = 60000;
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行第 " + Step + " 步校准出错: " + Environment.NewLine + ex.Message);
            }
        }


        public List<string> _4PortCalIni(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            try
            {
                List<string> StepInfor = new List<string>();
                ClearError();
                string _calkit = CiList[0].Calibration.CalKit;
                string _calmethod = CiList[0].Calibration.CalMethod;

                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    SelectChannel(chaninfor.Chan);
                    this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:CKIT " + CalKitLookUp[_calkit]);
                    this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:METH:" + CalMethodLookUp[_calmethod] + "4 1,2,3,4");
                });

                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port1");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port1");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port1");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " OPEN to Port4");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " SHORT to Port4");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " LOAD to Port4");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port1 and Port2");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port1 and Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port1 and Port4");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port2 and Port3");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port2 and Port4");
                StepInfor.Add("Connect " + calibrationInfor.CalKit + " TO " + calibrationInfor.CalKit + " ADAPTER between Port3 and Port4");
                MonitorSystem();
                return StepInfor;

            }
            catch (Exception ex)
            {
                throw new Exception("设置校准件信息出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void _4Port_Step(List<ChanInfor> CiList, Gender ConnectorGender, int Step)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    SelectChannel(chaninfor.Chan);
                    this.NaConnect.timeout = 300000;
                    switch (Step)
                    {
                        case 1:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 1");
                            break;
                        case 2:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 1");
                            break;
                        case 3:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 1");
                            break;
                        case 4:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 2");
                            break;
                        case 5:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 2");
                            break;
                        case 6:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 2");
                            break;
                        case 7:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 3");
                            break;
                        case 8:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 3");
                            break;
                        case 9:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 3");
                            break;
                        case 10:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:OPEN 4");
                            break;
                        case 11:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC " + (Int32)ConnectorGender);
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SHOR 4");
                            break;
                        case 12:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:LOAD 4");
                            break;
                        case 13:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 1,2");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 2,1");
                            break;
                        case 14:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 1,3");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 3,1");
                            break;
                        case 15:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 1,4");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 4,1");
                            break;
                        case 16:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 2,3");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 3,2");
                            break;
                        case 17:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 2,4");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 4,2");
                            break;
                        case 18:
                            this.NaConnect.timeout = 480000;
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:SUBC 1");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 3,4");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:THRU 4,3");
                            break;
                    }
                    this.NaConnect.timeout = 60000;
                });
                MonitorSystem();
            }
            catch (Exception ex)
            {
                throw new Exception("执行第 " + Step + " 步校准出错: " + Environment.NewLine + ex.Message);
            }
        }


        public void ESOLTStartCal(List<ChanInfor> CiList)
        {
            try
            {
                ClearError();
                CiList.ForEach(delegate(ChanInfor chaninfor)
                {
                    this.NaConnect.timeout = 480000;
                    switch (chaninfor.Calibration.CalMethod)
                    {
                        case "ESolt2PortCal":
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:ECAL:SOLT2 1,2");
                            break;

                        case "ETRANCal":
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:ECAL:THRU 1,2");
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:ECAL:THRU 2,1");
                            break;

                        default:
                            this.ReWriteNaConnect.Write(@":SENS" + chaninfor.Chan + ":CORR:COLL:ECAL:SOLT2 1,2");
                            break;
                    }
                    this.NaConnect.timeout = 60000;
                    MonitorSystem();

                });
            }
            catch (Exception ex)
            {
                throw new Exception("执行校准出错: " + Environment.NewLine + ex.Message);
            }
        }

        public void ESOLTCalFinish(List<ChanInfor> CiList)
        {
            //不需要
        }

        public void ESOLTCalAbort(List<ChanInfor> CiList)
        {
            //不需要
        }


        public ItemTestValues MeasMaxAndPhaseMemory(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                string trformat = "";
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":SEL:FORM?");
                trformat = this.NaConnect.readStr().Trim();
                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);
                SetTraceFormat(itemTestRequir.Ch, Format.PHAS);
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MATH:MEM");
                SetTraceFormat(itemTestRequir.Ch, (Format)Enum.Parse(typeof(Format), trformat));

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasMaxAndPhaseMemory出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasMinAndPhaseMemory(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                string trformat = "";
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":SEL:FORM?");
                trformat = this.NaConnect.readStr().Trim();
                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum(), 3);
                //读取频点
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:X?");
                XYValues.MeasuredFreq = System.Math.Round(this.NaConnect.readNum() / 1000000, 3);
                SetTraceFormat(itemTestRequir.Ch, Format.PHAS);
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MATH:MEM");
                SetTraceFormat(itemTestRequir.Ch, (Format)Enum.Parse(typeof(Format), trformat));

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasMinAndPhaseMemory出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasPhaseBalance(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            double maxvalue;
            double minvalue;
            try
            {

                if (testData == null)
                {
                    throw new Exception("执行MeasPhaseBalance出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }
                if (!testData.CheckTestIndexKey(Convert.ToInt32(itemTestRequir.Interval)))
                {
                    throw new Exception("执行MeasPhaseBalance出错：" + Environment.NewLine + "因为第" + itemTestRequir.Interval.ToString() + "步未测试,无法调用它的测试信息!");
                }

                string trformat = "";
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":SEL:FORM?");
                trformat = this.NaConnect.readStr().Trim();

                SetTraceFormat(itemTestRequir.Ch, Format.PHAS);
                //执行除法
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC DIV");


                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                maxvalue = System.Math.Round(this.NaConnect.readNum(), 3);

                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                minvalue = System.Math.Round(this.NaConnect.readNum(), 3);

                XYValues.MeasuredValue = Math.Round((maxvalue - minvalue), 3);
                XYValues.MeasuredFreq = 0;

                //还原
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC NORM");
                SetTraceFormat(itemTestRequir.Ch, (Format)Enum.Parse(typeof(Format), trformat));

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasPhaseBalance出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasRippleByDeviationFromLinearPhase(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值
                List<mkPiont> ValidValue = new List<mkPiont>();
                List<List<mkPiont>> PhaseValue = new List<List<mkPiont>>();
                List<List<mkPiont>> RippleValue = new List<List<mkPiont>>();
                double interval = itemTestRequir.Interval;
                double Gap = 0;
                mkPiont ripple = new mkPiont() { xPos = 0, yPos = 0 };
                double temp = 0;


                string trformat = "";
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":SEL:FORM?");
                trformat = this.NaConnect.readStr().Trim();
                SetTraceFormat(itemTestRequir.Ch, Format.PHAS);

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");
                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");
                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }

                for (int no = 0; no < ValidValue.Count - 1; no++)
                {
                    for (int j = no; j < ValidValue.Count - 1; j++)
                    {
                        Gap = ValidValue[no].yPos - ValidValue[j + 1].yPos;
                        if (Gap >= (interval - 0.1) && Gap <= (interval + 0.1))
                        {
                            PhaseValue.Add(new List<mkPiont>() { ValidValue[no], ValidValue[j + 1] });
                        }
                    }
                }

                if (PhaseValue.Count() <= 0)
                {
                    XYValues.MeasuredValue = Math.Round(itemTestInfor.HighSpec, 3);
                    XYValues.MeasuredFreq = 0;
                }
                else
                {

                    SetTraceFormat(itemTestRequir.Ch, Format.MLOG);
                    this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                    NAValueData = this.NaConnect.readStr().Split(',');
                    DataArray = new double[2, NAValueData.Length / 2];
                    ValidValue = new List<mkPiont>();

                    Dindex = 0;
                    for (Int32 position = 0; position < NAValueData.Length; position += 2)
                    {
                        DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                        Dindex += 1;
                    }

                    for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                    {
                        ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                    }

                    foreach (List<mkPiont> mkgroup in PhaseValue)
                    {
                        RippleValue.Add(new List<mkPiont>() {new mkPiont{ xPos = mkgroup[0].xPos, yPos = ValidValue.Find( value => value.xPos == mkgroup[0].xPos).yPos}
                    ,new mkPiont{ xPos = mkgroup[1].xPos,yPos = ValidValue.Find(value => value.xPos == mkgroup[1].xPos).yPos}});
                    }

                    foreach (List<mkPiont> ripgroup in RippleValue)
                    {
                        temp = ripgroup.Max(yvalue => yvalue.yPos) - ripgroup.Min(yvalue => yvalue.yPos);
                        if (temp > ripple.yPos)
                        {
                            ripple.yPos = temp;
                            ripple.xPos = ripgroup[0].xPos;
                        }
                    }

                    XYValues.MeasuredValue = Math.Round(ripple.yPos, 3);
                    XYValues.MeasuredFreq = Math.Round(ripple.xPos / 1000000, 3);
                }
                //还原
                SetTraceFormat(itemTestRequir.Ch, (Format)Enum.Parse(typeof(Format), trformat));

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasRippleByDeviationFromLinearPhase出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasAverageVar(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值
                List<mkPiont> ValidValue = new List<mkPiont>();

                double Average = 1000;
                double temp = 0;
                string interval = itemTestRequir.Interval.ToString();
                Int32 Gap = Int32.Parse(interval);

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");
                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");
                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }

                if (Gap < ValidValue.Count())
                {
                    List<mkPiont> DivideValidVal = new List<mkPiont>();
                    for (Int32 divide = 0; divide < (ValidValue.Count() - itemTestRequir.Interval); divide += 1)
                    {
                        DivideValidVal = ValidValue.GetRange(divide, Convert.ToInt32(itemTestRequir.Interval) + 1);
                        temp = DivideValidVal.Average(item => item.yPos);
                        if (temp < Average) Average = temp;
                    }

                }
                else
                {
                    Average = ValidValue.Average(item => item.yPos);
                }
                XYValues.MeasuredValue = Math.Round(Average, 3);
                XYValues.MeasuredFreq = 0;
                MonitorSystem();

                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasAverageVar出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToDataMemory(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {

                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;

                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MATH:MEM");
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasDelayAndMemory出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasDelayBalance(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            double maxvalue;
            double minvalue;
            double absmaxvalue;
            try
            {

                if (testData == null)
                {
                    throw new Exception("执行MeasDelayBalance出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }
                if (!testData.CheckTestIndexKey(Convert.ToInt32(itemTestRequir.Interval)))
                {
                    throw new Exception("执行MeasDelayBalance出错：" + Environment.NewLine + "因为第" + itemTestRequir.Interval.ToString() + "步未测试,无法调用它的测试信息!");
                }


                //执行除法
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC DIV");


                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Search Range
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STAR " + itemTestRequir.StarF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM:STOP " + itemTestRequir.StopF * 1000000);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:DOM ON");
                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MAX);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                maxvalue = System.Math.Round(Math.Abs(this.NaConnect.readNum() * 1000000000), 3);

                //设置Marker Search
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:TYPE " + mkType.MIN);
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:FUNC:EXEC");
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                minvalue = System.Math.Round(Math.Abs(this.NaConnect.readNum() * 1000000000), 3);

                absmaxvalue = new double[2] { maxvalue, minvalue }.Max();
                XYValues.MeasuredValue = Math.Round(absmaxvalue, 3);
                XYValues.MeasuredFreq = 0;

                //还原
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC NORM");


                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasDelayBalance出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasBandWidth(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues XYValues;
            try
            {
                this.NaConnect.timeout = 6000;
                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Marker点
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:X " + itemTestRequir.StarF * 1000000);
                //设置Bandwidth Value
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:BWID:THR " + itemTestRequir.Interval);
                //打开Bandwidth 状态
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK:BWID ON");

                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:BWID:DATA?");
                string[] data = this.NaConnect.readStr().Trim().Split(',');
                //关闭Bandwidth 状态
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK:BWID OFF");
                //Bw
                XYValues.MeasuredValue = System.Math.Round(Convert.ToDouble(data[0].Trim()) / 1000000, 3);
                //Cent
                XYValues.MeasuredFreq = System.Math.Round(Convert.ToDouble(data[1].Trim()) / 1000000, 3);
                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                ClearError();
                Console.Write(ex.Message);
                return new ItemTestValues() { MeasuredFreq = 0, MeasuredValue = 0 };
            }
        }

        public ItemTestValues MeasToAluDelay(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //打开Mark1
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:STAT ON");
                //设置Marker点
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MARK1:X " + itemTestRequir.StarF * 1000000);
                //读取值
                this.NaConnect.sendCmd(@":CALC" + itemTestRequir.Ch + ":MARK1:Y?");
                XYValues.MeasuredValue = System.Math.Round(this.NaConnect.readNum() * 1000000000, 3);
                XYValues.MeasuredFreq = System.Math.Round(itemTestRequir.StarF, 3);

                Channel channel = new Channel();
                channel.ChannelName = itemTestInfor.ItemName.Substring(itemTestInfor.ItemName.IndexOf(':') + 1, itemTestInfor.ItemName.IndexOf('[') - itemTestInfor.ItemName.IndexOf(':') - 1).Trim();
                channel.FreqStep = itemTestRequir.Interval;
                channel.StartToStopFreq = itemTestRequir.StarF.ToString("0.00") + "-" + itemTestRequir.StopF.ToString("0.00");
                List<ValueSpec> DelayValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();

                tempVal.MeasuredFreq = Math.Round(XYValues.MeasuredFreq, 2);
                tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                tempVal.MeasuredValue = XYValues.MeasuredValue;
                tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                DelayValues.Add(tempVal);

                channel.ChannelStatus = ResultStatus.Passed;
                foreach (ValueSpec val in DelayValues)
                {
                    if (val.result == ResultStatus.Failed)
                    {
                        channel.ChannelStatus = ResultStatus.Failed;
                        break;
                    }
                }
                channel.Values = DelayValues;

                if (testData != null)
                {
                    TempDatas.RecordData("Group Delay Test", channel);
                }

                MonitorSystem();
                return XYValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasToAluDelay出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToAluAvgIL(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值
                List<mkPiont> ValidValue = new List<mkPiont>();

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");
                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");
                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position], yPos = DataArray[0, position] });
                }

                double AvgData = ValidValue.Average(item => item.yPos);
                double AvgFreq = ValidValue.Average(item => item.xPos);


                Channel channel = new Channel();
                channel.ChannelName = itemTestInfor.ItemName.Substring(itemTestInfor.ItemName.IndexOf(':') + 1, itemTestInfor.ItemName.IndexOf('[') - itemTestInfor.ItemName.IndexOf(':') - 1).Trim();
                channel.FreqStep = itemTestRequir.Interval;
                channel.StartToStopFreq = itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0");
                List<ValueSpec> ILValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();

                tempVal.MeasuredFreq = Math.Round(AvgFreq / 1000000, 2);
                tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                tempVal.MeasuredValue = Math.Round(AvgData, 3);
                tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                ILValues.Add(tempVal);

                channel.ChannelStatus = ResultStatus.Passed;
                foreach (ValueSpec val in ILValues)
                {
                    if (val.result == ResultStatus.Failed)
                    {
                        channel.ChannelStatus = ResultStatus.Failed;
                        break;
                    }
                }
                channel.Values = ILValues;

                if (testData != null)
                {
                    TempDatas.RecordData("Average Insertion Loss", channel);
                }

                XYValues.MeasuredValue = Math.Round(AvgData, 3);
                XYValues.MeasuredFreq = Math.Round(AvgFreq / 1000000, 3);
                MonitorSystem();
                return XYValues;
            }


            catch (Exception ex)
            {
                throw new Exception("执行MeasToAluAvgIL出错：" + Environment.NewLine + ex.Message);
            }
        }


        public List<string> E8PortCalIni(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            throw new NotImplementedException();
        }

        public void E8PortStep(List<ChanInfor> CiList, Gender ConnectorGender, int Step)
        {
            throw new NotImplementedException();
        }

        public void E8PortCalFinish(List<ChanInfor> CiList, Gender ConnectorGender)
        {
            throw new NotImplementedException();
        }


        public List<string> E10PortCalIni(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            throw new NotImplementedException();
        }

        public List<string> E12PortCalIni(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            throw new NotImplementedException();
        }


        public List<string> E2PortCal9_10Ini(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasDelayCoherence(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                string[] IndexSteps = itemTestRequir.State.Split(',');
                if (testData == null)
                {
                    throw new Exception("执行MeasDelayCoherence出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }
                foreach (string strindex in IndexSteps)
                {
                    if (!testData.CheckTestIndexKey(Convert.ToInt32(strindex)))
                    {
                        throw new Exception("执行MeasDelayCoherence出错：" + Environment.NewLine + "第" + strindex + "步可能还未测试,无法调用它的测试信息!");
                    }
                }
                //先从TempExcelDatas中找到Group Delay 的 HWTempData
                HWTempData _TempData = TempExcelDatas.TempDatas.Find(delegate(HWTempData dhwTemp)
                {
                    return dhwTemp.TestItemName == "Group Delay";
                });

                if (_TempData == null)
                {
                    throw new Exception("执行MeasDelayCoherence出错：" + Environment.NewLine + "GroupDelay数据未采集,无法调用它的测试信息,请通知工程师!");
                }

                //再找出频率范围在设定范围内的channel

                var channels = from ch in _TempData.TestChanels
                               where ch.StartToStopFreq == itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0")
                               select ch;

                //再两两分组
                List<Channel> Lchannels = channels.ToList();
                int channelcount = Lchannels.Count();
                if (channelcount % 2 != 0)
                {
                    throw new Exception("执行MeasDelayCoherence出错：" + Environment.NewLine + "GroupDelay数据未成对采集,请通知工程师!");
                }

                int channelcouple = channelcount / 2;
                List<double> coupleDelay = new List<double>();
                List<List<double>> CoupleDalays = new List<List<double>>();
                double tempdelayval = 0;

                for (int i = 0; i < Lchannels[0].Values.Count; i++)
                {
                    coupleDelay = new List<double>();
                    for (int j = 0; j < channelcouple; j++)
                    {
                        coupleDelay.Add(Lchannels[j * 2 + 1].Values[i].MeasuredValue - Lchannels[j * 2].Values[i].MeasuredValue);
                    }
                    tempdelayval = coupleDelay.Max() - coupleDelay.Min();
                    if (tempdelayval > XYValues.MeasuredValue)
                    {
                        XYValues.MeasuredValue = tempdelayval;
                    }
                    //其实这一步已不需要做，前面步骤已经将值算出
                    CoupleDalays.Add(coupleDelay);
                }

                MonitorSystem();
                XYValues.MeasuredValue = Math.Round(XYValues.MeasuredValue, 3);
                return XYValues;


            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasDelayCoherence出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasToHuaWeiExcelPhase(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                //测试带宽起始频率相对状态文件起始频率的位置
                Int32 FreqPosL;
                Int32 FreqPosH;
                //获取网分X轴频率数据
                string[] NAFreqData;
                //网分X轴频率数据,由NAFreqData转化
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                //测试带宽内，频点与其对应数值

                List<mkPiont> ValidValue = new List<mkPiont>();
                //执行除法
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC DIV");

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));
                FreqPosL = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000);
                if (FreqPosL == -1) throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");

                FreqPosH = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000);
                if (FreqPosH == -1) throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');

                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }

                for (Int32 position = FreqPosL; position <= FreqPosH; position++)
                {
                    ValidValue.Add(new mkPiont() { xPos = Frequencies[position] / 1000000, yPos = DataArray[0, position] });
                }

                Channel channel = new Channel();
                channel.ChannelName = itemTestInfor.ItemName.Substring(itemTestInfor.ItemName.IndexOf(':') + 1, itemTestInfor.ItemName.IndexOf('[') - itemTestInfor.ItemName.IndexOf(':') - 1).Trim();
                channel.FreqStep = itemTestRequir.Interval;
                channel.StartToStopFreq = itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0");
                List<ValueSpec> PhaseValues = new List<ValueSpec>();
                ValueSpec tempVal = new ValueSpec();
                int stepcount = 0;
                for (double freq = itemTestRequir.StarF; Math.Round(freq, 1) <= itemTestRequir.StopF; freq += itemTestRequir.Interval)
                {
                    stepcount++;
                    if (stepcount == 2 && itemTestRequir.StarF % 1 != 0 && itemTestRequir.Interval % 1 == 0)
                    {
                        tempVal.MeasuredFreq = Math.Ceiling(itemTestRequir.StarF);
                        freq = Math.Ceiling(itemTestRequir.StarF);
                    }
                    tempVal.MeasuredFreq = Math.Round(freq, 1);
                    tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                    tempVal.MeasuredValue = Math.Round(SearchCharData(tempVal.MeasuredFreq, ValidValue), 2);
                    tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                    PhaseValues.Add(tempVal);

                    if (tempVal.MeasuredFreq + itemTestRequir.Interval > itemTestRequir.StopF && tempVal.MeasuredFreq != itemTestRequir.StopF)
                    {
                        tempVal.MeasuredFreq = itemTestRequir.StopF;
                        tempVal.Spc = itemTestInfor.LowSpec + ":" + itemTestInfor.HighSpec;
                        tempVal.MeasuredValue = Math.Round(SearchCharData(tempVal.MeasuredFreq, ValidValue), 2);
                        tempVal.result = (tempVal.MeasuredValue >= itemTestInfor.LowSpec && tempVal.MeasuredValue <= itemTestInfor.HighSpec ? ResultStatus.Passed : ResultStatus.Failed);
                        PhaseValues.Add(tempVal);
                    }
                }
                channel.ChannelStatus = ResultStatus.Passed;
                foreach (ValueSpec val in PhaseValues)
                {
                    if (val.result == ResultStatus.Failed)
                    {
                        channel.ChannelStatus = ResultStatus.Failed;
                        break;
                    }
                }
                channel.Values = PhaseValues;

                if (testData != null)
                {
                    TempExcelDatas.RecordData("Phase", channel);
                }

                //还原
                this.ReWriteNaConnect.Write(@":CALC" + itemTestRequir.Ch + ":MATH:FUNC NORM");

                XYValues.MeasuredValue = PhaseValues.Max(val => val.MeasuredValue);
                XYValues.MeasuredFreq = PhaseValues.First(val => val.MeasuredValue == XYValues.MeasuredValue).MeasuredFreq;
                MonitorSystem();
                return XYValues;
            }
            catch (FormatException FEx)
            {
                throw new Exception("执行MeasToHuaWeiExcelPhase出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + FEx.Message);
            }

            catch (OverflowException OEx)
            {
                throw new Exception("执行MeasToHuaWeiExcelPhase出错: " + Environment.NewLine + "请检查指标设置是否合理！" + Environment.NewLine + OEx.Message);
            }

            catch (Exception ex)
            {
                throw new Exception("执行MeasToHuaWeiExcelPhase出错：" + Environment.NewLine + ex.Message);
            }
        }

        public ItemTestValues MeasPhaseCoherence(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            try
            {
                ItemTestValues XYValues;
                XYValues.MeasuredValue = 0;
                XYValues.MeasuredFreq = 0;
                string[] IndexSteps = itemTestRequir.State.Split(',');
                if (testData == null)
                {
                    throw new Exception("执行MeasPhaseCoherence出错：" + Environment.NewLine + "此步骤不可以进行选择性测试!");
                }
                foreach (string strindex in IndexSteps)
                {
                    if (!testData.CheckTestIndexKey(Convert.ToInt32(strindex)))
                    {
                        throw new Exception("执行MeasPhaseCoherence出错：" + Environment.NewLine + "第" + strindex + "步可能还未测试,无法调用它的测试信息!");
                    }
                }
                //先从TempExcelDatas中找到Group Delay 的 HWTempData
                HWTempData _TempData = TempExcelDatas.TempDatas.Find(delegate(HWTempData dhwTemp)
                {
                    return dhwTemp.TestItemName == "Phase";
                });

                if (_TempData == null)
                {
                    throw new Exception("执行MeasPhaseCoherence出错：" + Environment.NewLine + "Phase数据未采集,无法调用它的测试信息,请通知工程师!");
                }

                //再找出频率范围在设定范围内的channel

                var channels = from ch in _TempData.TestChanels
                               where ch.StartToStopFreq == itemTestRequir.StarF.ToString("0.0") + "-" + itemTestRequir.StopF.ToString("0.0")
                               select ch;

                //再两两分组
                List<Channel> Lchannels = channels.ToList();
                int channelcount = Lchannels.Count();
                //此计算方式不需要再调用两个结果的差值2015-8-9
                int channelcouple = channelcount;
                List<double> coupleDelay = new List<double>();
                List<List<double>> CoupleDalays = new List<List<double>>();
                double tempdelayval = 0;
                double tmp = 0;

                for (int i = 0; i < Lchannels[0].Values.Count; i++)
                {
                    coupleDelay = new List<double>();
                    for (int j = 0; j < channelcouple; j++)
                    {
                        tmp = Lchannels[j].Values[i].MeasuredValue;
                        coupleDelay.Add(tmp);
                    }
                    tempdelayval = coupleDelay.Max() - coupleDelay.Min();
                    if (tempdelayval > 180)
                    {
                        tempdelayval = 360 - tempdelayval;
                    }
                    if (tempdelayval > XYValues.MeasuredValue)
                    {
                        XYValues.MeasuredValue = tempdelayval;
                    }
                    //其实这一步已不需要做，前面步骤已经将值算出
                    CoupleDalays.Add(coupleDelay);
                }

                MonitorSystem();
                return XYValues;


            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasPhaseCoherence出错：" + Environment.NewLine + ex.Message);
            }
        }


        public List<string> E2PortCal1_2Ini(List<ChanInfor> CiList, CalibrationInfor calibrationInfor)
        {
            throw new NotImplementedException();
        }


        public void SetContinuous(int activech)
        {
            throw new NotImplementedException();
        }

        public void SetHold(int activech)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasToLimitJPGS2P(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasSubtract(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }





        public ItemTestValues MeasMemSave(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }

        public ItemTestValues MeasPhaseCompareA(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }

        public ItemTestValues MeasPhaseCompareB(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasBandWidthLeft(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }

        public ItemTestValues MeasBandWidthRight(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasTargetWidthLeft(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }

        public ItemTestValues MeasTargetWidthRight(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasToS3P(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }


        public ItemTestValues MeasAveragePower(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }

        public ItemTestValues MeasAverageVarPower(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            ItemTestValues result;
            try
            {
                //获取网分X轴频率数据
                string[] NAFreqData;
                List<double> Frequencies = new List<double>();
                //网分Y轴数值
                string[] NAValueData;
                //处理NAValueData
                double[,] DataArray = null;
                List<mkPiont> list = new List<mkPiont>();
                double num = 1000.0;
                string s = itemTestRequir.Interval.ToString(CultureInfo.InvariantCulture);
                int num2 = int.Parse(s);

                this.NaConnect.sendCmd(@":SENSE" + itemTestRequir.Ch + ":FREQ:DATA?");
                NAFreqData = this.NaConnect.readStr().Split(',');
                Array.ForEach(NAFreqData, Freq => Frequencies.Add(Convert.ToDouble(Freq)));

                //开始位置
                int num3 = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StarF * 1000000.0);
                if (num3 == -1)
                {
                    throw new Exception("未能找到TestSequece中设置的起始频率点！请检查指标设置");
                }
                //结束位置
                int num4 = Frequencies.FindIndex(Freq => Freq == itemTestRequir.StopF * 1000000.0);
                if (num4 == -1)
                {
                    throw new Exception("未能找到TestSequece中设置的终止频率点！请检查指标设置");
                }

                this.NaConnect.sendCmd(@":FORM:DATA ASC;:CALC" + itemTestRequir.Ch + ":DATA:FDAT?");
                NAValueData = this.NaConnect.readStr().Split(',');
                DataArray = new double[2, NAValueData.Length / 2];

                Int32 Dindex = 0;
                for (Int32 position = 0; position < NAValueData.Length; position += 2)
                {
                    DataArray[0, Dindex] = Convert.ToDouble(NAValueData[position]);
                    Dindex += 1;
                }
                //获取所有值
                for (int i = num3; i <= num4; i++)
                {
                    list.Add(new mkPiont
                    {
                        xPos = Frequencies[i],
                        yPos = DataArray[0, i]
                    });
                }
                //var lt = list.Select(p => Math.Pow(10, p.yPos/10));
                //double num3 = Math.Log10(lt.Average())*10;
                //double value = Math.Log10(lt.Average()) * 10;
                if (num2 < list.Count())
                {
                    List<mkPiont> source;
                    int num5 = 0;
                    while (num5 < list.Count() - itemTestRequir.Interval)
                    {
                        source = list.GetRange(num5, Convert.ToInt32(itemTestRequir.Interval) + 1);
                        var lt = source.Select(p => Math.Pow(10, p.yPos / 10));
                        double num6 = Math.Log10(lt.Average()) * 10;
                        if (num6 < num)
                        {
                            num = num6;
                        }
                        num5++;
                    }
                }
                else
                {
                    var lt = list.Select(p => Math.Pow(10, p.yPos / 10)).ToList();
                    num = Math.Log10(lt.Average()) * 10;
                }
                ItemTestValues itemTestValues;
                itemTestValues.MeasuredValue = Math.Round(num, 3);
                itemTestValues.MeasuredFreq = 0.0;
                this.MonitorSystem();
                result = itemTestValues;
            }
            catch (Exception ex)
            {
                throw new Exception("执行MeasAverageVarPower出错：" + Environment.NewLine + ex.Message);
            }
            return result;
        }
        public ItemTestValues MeasMultipRej(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }
        public ItemTestValues MeasPhaseCompareC(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }
        public ItemTestValues MeasPhaseCompareD(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }
        public ItemTestValues MeasSmith(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }
        public ItemTestValues MeasToSmithS2P(ItemTestRequir itemTestRequir, ItemTestInfor itemTestInfor, TestData testData = null, HWTempDatas TempDatas = null, HWTempDatas TempExcelDatas = null)
        {
            throw new NotImplementedException();
        }
    }
}
