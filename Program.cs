using Aspose.Cells;

class TestClass
{
    static void Main()
    {
        Workbook wb = new Workbook("data.xlsx");
        WorksheetCollection collection = wb.Worksheets;
        for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
        {

            Worksheet worksheet = collection[worksheetIndex];
            Console.WriteLine("Enter initial power:");
            double power = Convert.ToDouble(Console.ReadLine());

            Line soundLine = new Line(worksheet, power);

            double result = soundLine.Evaluation();
            Console.WriteLine(result);
        }
    }
}

public class Line
{
    public double Power { get; set; }
    public int Number { get; set; }

    public double[] DecayCoeff { get; set; }

    public double[] GainCoeff { get; set; }

    public double[] NoiseCoeff { get; set; }

    double noicePower = 0;
    public Line(Worksheet worksheet, double power)
    {
        Power= Math.Pow(10, 0.1 * (power - 30)) ; // in dB

        int rows = worksheet.Cells.MaxDataRow;

        Number = rows+1;

        DecayCoeff = new double[Number];
        GainCoeff = new double[Number];
        NoiseCoeff = new double[Number];

        Filling(DecayCoeff, worksheet, 0, x => Math.Pow(10, 0.1 * x ) );
        Filling(GainCoeff, worksheet, 1, x => Math.Pow(10, 0.1 * x ));
        Filling(NoiseCoeff, worksheet, 2, x => Math.Pow(10, 0.1 * ( x - 30) ));  
    }

    public double Evaluation()
    {
        
        for(int i = 0; i < Number; i++)
        {
            Power = Power + GainCoeff[i] - DecayCoeff[i];
            noicePower = noicePower + NoiseCoeff[i] + GainCoeff[i] - DecayCoeff[i];
        }
        return 10*Math.Log10(Power - noicePower);
    }

    //i is number of column
    void Filling(double[] data, Worksheet worksheet, int i, Func<double, double> func )
    {
        for (int j = 0; j < Number; j++)
        {
            var value = func((double)worksheet.Cells[j, i].Value);
            data[j] = value;
        }
    }


}