using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using LiveCharts.Geared;


namespace Dashboard
{
    public class SpeedTestVm
    {
        public double _trend = 0;


        public SpeedTestVm()
        {
            Values = new GearedValues<double>().WithQuality(Quality.High);
        }

        public bool IsReading { get; set; }
        public GearedValues<double> Values { get; set; }
        public double Count { get; set; }
        public double CurrentLecture { get; set; }
        public bool IsHot { get; set; }

        public void Stop()
        {
            IsReading = false;
        }

        public void Clear()
        {
            Values.Clear();
        }

        public void Read()
        {
            if (IsReading) return;

            //lets keep in memory only the last 20000 records,
            //to keep everything running faster
            const int keepRecords = 450;
            IsReading = true;

            Action readFromTread = () =>
            {
                while (IsReading)
                {
                    Thread.Sleep(50);
                    var first = Values.DefaultIfEmpty(0).FirstOrDefault();
                    if (Values.Count > keepRecords - 1) Values.Remove(first);
                    if (Values.Count < keepRecords) Values.Add(_trend);
                    IsHot = _trend > 0;
                    Count = Values.Count;
                    CurrentLecture = _trend;
                }
            };

            //2 different tasks adding a value every ms
            //add as many tasks as you want to test this feature
            Task.Factory.StartNew(readFromTread);
            Task.Factory.StartNew(readFromTread);
            //Task.Factory.StartNew(readFromTread);
            //Task.Factory.StartNew(readFromTread);
            //Task.Factory.StartNew(readFromTread);
            //Task.Factory.StartNew(readFromTread);
            //Task.Factory.StartNew(readFromTread);
            //Task.Factory.StartNew(readFromTread);
        }
    }
}