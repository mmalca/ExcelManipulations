#include "Spire.Xls.o.h"
#include "isr84lib.cpp"

using namespace Spire::Xls;
//using namespace Spire::Common;
using namespace std;

int main()
{
  
    // Initialize an instance of the Workbook class
    intrusive_ptr<Workbook> workbook = new Workbook();

    try {
        //Load an XLSX or XLS file
        workbook->LoadFromFile(L"C:/Users/malcabo/OneDrive - huji.ac.il/Specify/Collections/Scorpions/Scorpion-TestLatLon.xlsx");
    }
    catch (...) {
        cerr << "Verify the file that you are trying to open exists on the path and is not in use..." << endl;
    }


    //Get the first worksheet (by default, a newly created workbook has 3 worksheets)
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
    
    //Get row count
    int rowCount = sheet->GetRows()->GetCount();

    int colLonITM = 53;
    int colLatITM = 54;
    
    int colLonNewWGS = 55;
    int colLatNewWGS = 56;

    int currentLonITM = 0;
    int currentLatITM = 0;

    double currentLonWGS = 0;
    double currentLatWGS = 0;

    ofstream logFile("outputLog.log");

   // if (logFile.is_open()) {
        // Redirect cout to the log file
        streambuf* coutbuf = cout.rdbuf();
        cout.rdbuf(logFile.rdbuf());
   // }
    //else {
        //cerr << "Unable to open the log file." << endl;
    //}

    for (int row = 2; row <= rowCount; row++)
    {

        if (currentLatITM = sheet->GetRange(row, colLatITM)->GetIsBlank()) {
            cout << "No ITM information at row number " << row << endl;
            continue;
        }

        cout << "Reading row: " << row << " --> " ;
        currentLonITM = sheet->GetRange(row, colLonITM)->GetNumberValue();
        currentLatITM = sheet->GetRange(row, colLatITM)->GetNumberValue();
        cout << "currentLonITM = " << currentLonITM << " ,  currentLatITM = " << currentLatITM << " -->  ";
        cout << "DONE" << endl;

        cout << "converting to WGS --> ";
        //convert to wgs
        itm2wgs84(currentLatITM, currentLonITM, currentLatWGS, currentLonWGS);
        cout << "currentLatWGS = " << currentLatWGS << " , " << "currentLonWGS = " << currentLonWGS << endl;

        cout << "saving data to file at: " << row << " , " << colLonNewWGS << " --> ";
        sheet->GetRange(row, colLonNewWGS)->SetNumberValue(currentLonWGS);
        sheet->GetRange(row, colLatNewWGS)->SetNumberValue(currentLatWGS);
        cout << "DONE" << endl;

    }
    // Restore cout
    std::cout.rdbuf(coutbuf);

    // Close the log file
    logFile.close();

    //save the updated excel file
    workbook->Save();
    workbook->Dispose();
    workbook.reset();

}