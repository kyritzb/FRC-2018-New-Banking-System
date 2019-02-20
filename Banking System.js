//**| by Bryan Kyritz team6969 |**\\
//**| Google Sheets Banking System|**\\
/*--------------------
|  Column            |
|  1 - TimeStamp     |
|  2 - Names         |
|  3 - Item          |
|  4 - Price         |
|  5 - Part/item     |
|  6 - Date          |
|  7 - Team          |
|  8 - Robot Area    |
|  9 - Store         |
//------------------*/
//--------------------------------------------------------------------------------------------------------------------------
//The function thats called when someone submits a form
//---------------------------------------------------------------------------------------------------------------------------
function OnSubmit(e) 
{
  updateIndividual();
  updateTeamSpending();
  updateBank();
}
//--------------------------------------------------------------------------------------------------------------------------
//Updates the bank
//---------------------------------------------------------------------------------------------------------------------------
function updateBank()
{
  var app = SpreadsheetApp;
  var logSheet = app.getActiveSpreadsheet().getSheetByName("Logs");
  var bank = app.getActiveSpreadsheet().getSheetByName("Bank Account");
  var people = [];
  var costs = [];
  var items = [];
  var whoCol = 6;
  var onWhatCol = 7;
  var nameCol = 2;
  var costCol = 8;
  var itemCol = 3;
  var priceCol = 4;
  var nameCol = 2;
  var total =0;
  var remainingCol = 11;
  var remainingRow = 5;
  var raisedCol = 4;
  var raisedRow = 2;
//gets all people and adds them to arrays------------------------------------
  for(var i =2; i<1000;i++)
  {
    bank.getRange(i+3,whoCol).setValue(logSheet.getRange(i, nameCol).getValue());
    bank.getRange(i+3,onWhatCol).setValue(logSheet.getRange(i, itemCol).getValue());
    bank.getRange(i+3,costCol).setValue(logSheet.getRange(i, priceCol).getValue());
    total+= logSheet.getRange(i, priceCol).getValue();
    if(logSheet.getRange(i, nameCol).isBlank())//if done with looking for people
    {
      break;
    }
  }
  bank.getRange(2,costCol).setValue(total);
  bank.getRange(remainingRow,remainingCol).setValue(bank.getRange(raisedRow, raisedCol).getValue() - total);
}
//--------------------------------------------------------------------------------------------------------------------------
//Updates the team Spending page
//---------------------------------------------------------------------------------------------------------------------------

function updateTeamSpending()
{
  var app = SpreadsheetApp;
  var logSheet = app.getActiveSpreadsheet().getSheetByName("Logs");
  var teamSpend = app.getActiveSpreadsheet().getSheetByName("Teams Spending");
  var nameCol = 2;
  var teamCol = 7;
  var itemCol = 3;
  var priceCol = 4;
  var startCol = 1;
  var startRow = 2;
  var mechItems =[];
  var elecItems=[];
  var artItems=[];
  var mechCosts=[];
  var elecCosts=[];
  var artCosts=[];
  var mechTotal = 0;
  var elecTotal = 0;
  var artTotal = 0;
 
  //------------------------------------------------------
  teamSpend.clear();
  for(var i = 2; i<1000; i++) //puts values into arrays
  {
    if(logSheet.getRange(i, teamCol).getValue() === "Mechanical")
    {
      mechItems.push(logSheet.getRange(i, itemCol).getValue());
      mechCosts.push(logSheet.getRange(i, priceCol).getValue());
    }
    else if(logSheet.getRange(i, teamCol).getValue() === "Electrical")
    {
      elecItems.push(logSheet.getRange(i, itemCol).getValue());
      elecCosts.push(logSheet.getRange(i, priceCol).getValue());
    }
    else if(logSheet.getRange(i, teamCol).getValue() === "Art")
    {
      artItems.push(logSheet.getRange(i, itemCol).getValue());
      artCosts.push(logSheet.getRange(i, priceCol).getValue());
    }
    else
    {
      break;
    }
  }
  
  for(var i = 0; i<mechItems.length; i++) //puts arrays of items into sheet
  {
    teamSpend.getRange(i+startRow, startCol).setValue(mechItems[i]);
    teamSpend.getRange(i+startRow, startCol+1).setValue(mechCosts[i]);
    mechTotal+=mechCosts[i];
  }
  for(var i = 0; i<elecItems.length; i++)
  {
    teamSpend.getRange(i+startRow, startCol+3).setValue(elecItems[i]);
    teamSpend.getRange(i+startRow, startCol+3+1).setValue(elecCosts[i]);
    elecTotal+=elecCosts[i];
  }
  for(var i = 0; i<artItems.length; i++)
  {
    teamSpend.getRange(i+startRow, startCol+6).setValue(artItems[i]);
    teamSpend.getRange(i+startRow, startCol+6+1).setValue(artCosts[i]);
    artTotal+=artCosts[i];
  }
  
  for(var i = 1; i<100; i++)
  {
    teamSpend.getRange(i, startCol+2).setBackgroundRGB(175, 175, 175);
    teamSpend.getRange(i, startCol+5).setBackgroundRGB(175, 175, 175);
    teamSpend.getRange(i, startCol+8).setBackgroundRGB(175, 175, 175);
  }
  teamSpend.getRange(startRow, startCol).setBackgroundRGB(255, 0, 0);
  teamSpend.getRange(startRow, startCol).setValue("Mechanical");
  teamSpend.getRange(startRow, startCol+1).setValue(mechTotal);
  teamSpend.getRange(startRow-1, startCol+1).setValue("Total");
  
  teamSpend.getRange(startRow, startCol+3).setBackgroundRGB(255, 255, 0);
  teamSpend.getRange(startRow, startCol+3).setValue("Electrical");
  teamSpend.getRange(startRow, startCol+4).setValue(elecTotal);
  teamSpend.getRange(startRow-1, startCol+4).setValue("Total");
  
  teamSpend.getRange(startRow, startCol+6).setBackgroundRGB(236, 80, 244);
  teamSpend.getRange(startRow, startCol+6).setValue("Art");
  teamSpend.getRange(startRow, startCol+7).setValue(artTotal);
  teamSpend.getRange(startRow-1, startCol+7).setValue("Total");
  
  console.log("mech Array:", mechItems.toString());
}

//--------------------------------------------------------------------------------------------------------------------------
//Updates the Individual Page
//---------------------------------------------------------------------------------------------------------------------------
function updateIndividual()
{  
  var app = SpreadsheetApp;
  var logSheet = app.getActiveSpreadsheet().getSheetByName("Logs");
  var IndiPurch = app.getActiveSpreadsheet().getSheetByName("Individual Purchases");
  var found = false;
  var nameCol = 2;
  var priceCol = 4;
  var people = [];
  var costs = [];
  var rowToStart = 3;
  var total = 0;
  var colToStart = 2;
  //Set up----------------------------------------------------------------
  IndiPurch.clear();
  IndiPurch.getRange(1, 1).setValue("Totals:");
  IndiPurch.getRange(3, 1).setValue("Names:");
  IndiPurch.getRange(4, 1).setValue("Puchases:");
  //gets all people and adds them to arrays------------------------------------
  for(var i =2; i<1000;i++)
  {
    found = false;
    for(var m = 0; m<people.length; m++)//searches through person array to find out if it needs to add the person
    {
      
      if(people[m] === logSheet.getRange(i, nameCol).getValue())
      {
        found = true;
      }
    }
    if(!found) //adds new person
    {
      people.push(logSheet.getRange(i,nameCol).getValue());
    }
    if(logSheet.getRange(i, nameCol).isBlank())//if done with looking for people
    {
      break;
    }
  }
  if(people.length>0)
  {
    //Gets people from array and writes them------------------------------------
    for(var i = 0; i<people.length; i++)
    {
      IndiPurch.getRange(rowToStart,i+colToStart).setValue(people[i]);
    }
    //gets array of costs for each person
    for(var p = 0; p<people.length; p++)
    {
      total = 0;
      costs = []; //reset array
      for(var i = 2; i< 1000; i++)//adds costs to "costs" array
      {
        if(logSheet.getRange(i, priceCol).isBlank())
        {
          break;
        }
        if(people[p] === logSheet.getRange(i,nameCol).getValue())
        {
          costs.push(logSheet.getRange(i, priceCol).getValue());
        }
      }
      for(var i = 0; i<costs.length; i++)//writes costs under name
      {
        IndiPurch.getRange(rowToStart+i + 1,p+colToStart).setValue(costs[i]);
      }
      for(var i = 0; i<costs.length;i++)//calculates total
      {
        total+=costs[i];
      }
      IndiPurch.getRange(rowToStart-2,p+colToStart).setValue(total); //writes total
    }
    IndiPurch.getRange(rowToStart-2,people.length+1).clear(); //fixes annoying number
    for(var i = 1; i<50; i++) //draws side lines
    {
      IndiPurch.getRange(i, people.length +1).setBackgroundRGB(175, 175, 175);
      IndiPurch.getRange(i, 1).setBackgroundRGB(175, 175, 175);
    }
    for(var i = 1; i<people.length;i++)//draws orange top line and total line
    {
      IndiPurch.getRange(rowToStart-1, i+1).setBackgroundRGB(244, 143, 66);
      IndiPurch.getRange(rowToStart-2, i+1).setBackgroundRGB(241, 249, 82);
    }
  }
}
