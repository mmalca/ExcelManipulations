#include "Spire.Xls.o.h"
#include "isr84lib.cpp"


using namespace Spire::Xls;
//using namespace Spire::Common;
using namespace std;

int main()
{

    ifstream file;
    string line;
    string word;
    vector<string> row = { 0 };

    int colLonITM = 53;
    int colLatITM = 54;

    int currentLonITM = 0;
    int currentLatITM = 0;

    double currentLatWGS = 0;
    double currentLonWGS = 0;

    file.open(L"C:/Users/malcabo/OneDrive - huji.ac.il/Specify/Collections/Scorpions/Scorpion-TestLatLon.xlsx");

    if (file.is_open()) {
        while (getline(file, line)) {
            stringstream str(line);
            while (getline(str, word, ','))
                row.push_back(word);
            //reading the current latlon
            currentLonITM = stoi(row[colLonITM]);
            currentLatITM = stoi(row[colLatITM]);

            if (currentLatITM == NULL || currentLonITM == NULL)
                continue;

            //convert to wgs
            itm2wgs84(currentLonITM, currentLatITM, currentLatWGS, currentLonWGS);

            //putting into the csv file the updated WGS
            ////..............here: need to put the calculated wgs - lan and lon at the correct place(column) and check if the program working////

        }
    }
    else
        cout << "Error openning file!" << endl;
    
    file.close();
    
    return 0;
    
    
    /*
    //Initialize an instance of the Workbook class
    intrusive_ptr<Workbook> workbook = new Workbook();
    //Get the first worksheet (by default, a newly created workbook has 3 worksheets)
    intrusive_ptr<Worksheet> sheet = dynamic_pointer_cast<Worksheet>(workbook->GetWorksheets()->Get(0));
    //Load an XLSX or XLS file
    workbook->LoadFromFile(L"C:/Users/malcabo/OneDrive - huji.ac.il/Specify/Collections/Scorpions/Scorpion-TestLatLon.xlsx");

    //Get row count
    int rowCount = sheet->GetRows()->GetCount();
    
    int colLonITM = 53;
    int colLatITM = 54;

    int colLonnewWGS = 0;
    int colLatnewWGS = 0;


    for (int row = 2; row < rowCount; row++) {
        //read lat and lon of itm
        int currentLonITM = sheet->GetRange(colLonITM, row)->GetHasNumber();
        int currentLatITM = sheet->GetRange(colLatITM, row)->GetHasNumber();

        cout << "currentLonITM = " << currentLonITM << endl;
        cout << "currentLatITM = " << currentLatITM << endl;
        break;

        if (currentLatITM == NULL || currentLonITM == NULL)
            continue;

        double currentLatWGS = 0;
        double currentLonWGS = 0;

        //convert to wgs
        itm2wgs84(currentLonITM, currentLatITM, currentLatWGS, currentLonWGS);

        //write to new lat lon columns
        sheet->GetRange(colLonnewWGS, row)->SetNumberValue(currentLonWGS);
        sheet->GetRange(colLatnewWGS, row)->SetNumberValue(currentLatWGS);

    }




    /*add data to cells
        sheet->GetRange(1, 3)->SetText(L"Salary");
        sheet->GetRange(2, 3)->SetNumberValue(6100);
    */

}