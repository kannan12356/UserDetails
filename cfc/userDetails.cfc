component{

    public function getData(){
        var userDetails = queryExecute("Select FirstName, LastName, Address, Email, Phone, DOB, Role From Users",[]);

        return userDetails;
    }

    remote function downloadPlainTemplate(){
        var mySheet = SpreadsheetNew("UserDetails",true);
        SpreadSheetAddRow(mySheet,"First Name,Last Name,Address,Email,Phone,DOB,Role");
        SpreadsheetFormatRow(mySheet, {'bold' : 'true'}, 1);

        cfheader( name="Content-Disposition", value="inline;filename=Plain_Template.xlsx" );
        cfcontent( variable=SpreadSheetReadBinary(mySheet), type="application/vnd.ms-excel" );
    }

    remote function downloadData(){
        var userDetails = getData();

        var mySheet = SpreadsheetNew("UserDetails",true);
        SpreadSheetAddRow(mySheet,"First Name,Last Name,Address,Email,Phone,DOB,Role");
        SpreadsheetFormatRow(mySheet, {'bold' : 'true'}, 1);
        SpreadSheetAddRows(mySheet,userDetails);

        cfheader( name="Content-Disposition", value="inline;filename=Template_with_data.xlsx" );
        cfcontent( variable=SpreadSheetReadBinary(mySheet), type="application/vnd.ms-excel" );
    }

    public function addData(required file){
        cffile( fileField="form.file", nameconflict="overwrite", destination="E:\ColdFusion\cfusion\wwwroot\UserDetails\Excel", action="upload", result="uploadFile" );
        var filePath = "#uploadFile.serverDirectory#\#uploadFile.clientFile#";

        cfspreadsheet( action="read", src="#filePath#", excludeheaderrow="true", headerrow=1, query="excelData" );

        cfoutput( query="excelData" ){
            
            if(Email != ""){
                var fName = excelData["First Name"];
                var lName = excelData["Last Name"];
                var address = Address;
                var email = Email;
                var phone = Phone;
                var DOB = DOB;
                var Role = Role;

                var emailCheck = queryExecute("Select * from Users Where Email = :email",
                {
                    email={cfsqltype:"cf_sql_nvarchar", value:email}
                }, {result="emailCheckResult"});

                var recordCount = emailCheckResult.recordCount;

                if(recordCount == 0){
                    var result = structNew();

                    if(fName == ""){
                        result.status = "First name is missing";
                    }
                    else if(lName == ""){
                        result.status = "Last name is missing";
                    }
                    else if(address == ""){
                        result.status = "address is missing";
                    }
                    else if(phone == ""){
                        result.status = "address is missing";
                    }
                    else if(DOB == ""){
                        result.status = "DOB is missing";
                    }
                    else if(Role == ""){
                        result.status = "Role is missing";
                    }
                    else{
                        result.status = "Success";
                    }

                    var insertData = queryExecute("INSERT into Users 
                    (FirstName, LastName, Address, Email, Phone, DOB, Role) 
                    Values 
                    (:fName, :lName, :address, :email, :phone, :DOB, :Role)",
                    {
                        fName={cfsqltype:"cf_sql_nvarchar", value:fName},
                        lName={cfsqltype:"cf_sql_nvarchar", value:lName},
                        address={cfsqltype:"cf_sql_nvarchar", value:address},
                        email={cfsqltype:"cf_sql_nvarchar", value:email},
                        phone={cfsqltype:"cf_sql_nvarchar", value:phone},
                        DOB={cfsqltype:"cf_sql_nvarchar", value:DOB},
                        Role={cfsqltype:"cf_sql_nvarchar", value:Role},
                    }, {result="addUserResult"});
                }
                else{
                    
                }
            }
        }
    }

}