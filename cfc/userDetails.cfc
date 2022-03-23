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

    private function arrayToQuery(data) {
        return data.reduce(function(accumulator, element) {
            element.each(function(key) {
                if (!accumulator.keyExists(key)) {
                    accumulator.addColumn(key, []);
                }
            });
            accumulator.addRow(element);
            return accumulator;
        }, QueryNew(""));
    }

    
    public function addData(required file){
        cffile( fileField="file", nameconflict="overwrite", destination="E:\ColdFusion\cfusion\wwwroot\UserDetails\Excel", action="upload", result="uploadFile" );
        var filePath = "#uploadFile.serverDirectory#\#uploadFile.clientFile#";

        cfspreadsheet( action="read", src="#filePath#", excludeheaderrow="true", headerrow=1, query="excelData" );

        var data = arrayNew(1);
        
        cfoutput( query="excelData" ){
            var userDetail = structNew();
            var result = arrayNew(1);

            var FirstName = excelData["First Name"];
            var LastName = excelData["Last Name"];
            var Address = excelData["Address"];
            var Email = excelData["Email"];
            var Phone = excelData["Phone"];
            var DOB = excelData["DOB"];
            DOB = dateFormat(DOB, "YYYY-mm-dd");
            var Role = excelData["Role"];

            if(FirstName != "" OR LastName != "" OR Address != "" OR Email != "" OR Phone != "" OR DOB != "" OR Role != ""){
                userDetail.FirstName = FirstName;
                userDetail.LastName = LastName;
                userDetail.Address = Address;
                userDetail.Email = Email;
                userDetail.Phone = Phone;
                userDetail.DOB = dateFormat(DOB, "dd-mm-YYYY");
                userDetail.Role = Role;
                
                if(Email != ""){
                    var emailCheck = queryExecute("Select * from Users Where Email = :email",
                    {
                        email={cfsqltype:"cf_sql_nvarchar", value:email}
                    }, {result="emailCheckResult"});

                    var recordCount = emailCheckResult.recordCount;
                    
                    if(recordCount == 0){

                        if(FirstName == ""){
                            var fNameErrorAdd = arrayAppend(result, "First name is missing");
                        }
                        if(LastName == ""){
                            var lNameErrorAdd = arrayAppend(result, "Last name is missing");
                        }
                        if(Address == ""){
                            var addressErrorAdd = arrayAppend(result, "address is missing");
                        }
                        if(Phone == ""){
                            var phoneErrorAdd = arrayAppend(result, "phone is missing");
                        }
                        if(DOB == ""){
                            var DOBErrorAdd = arrayAppend(result, "DOB is missing");
                        }
                        if(Role == ""){
                            var RoleErrorAdd = arrayAppend(result, "Role is missing");
                        }
                        if(FirstName != "" AND LastName != "" AND Address != "" AND Phone != "" AND DOB != "" AND Role != ""){
                            var statusErrorAdd = arrayAppend(result, "Success");

                            var RoleArray = listToArray(Role);

                            var Roles = arrayNew(1);

                            for(i=1; i<=arrayLen(RoleArray); i++){
                                var roleCheck = queryExecute("Select * from Roles Where Role = :role",
                                {
                                    role={cfsqltype:"cf_sql_nvarchar", value:RoleArray[i]}
                                }, {result="roleCheckResult"});

                                if(roleCheckResult.recordCount == 1){
                                    arrayAppend(Roles, RoleArray[i]);
                                }
                            }
                            
                            Role = arrayToList(Roles, ",");

                            var insertData = queryExecute("INSERT into Users 
                            (FirstName, LastName, Address, Email, Phone, DOB, Role) 
                            Values 
                            (:fName, :lName, :address, :email, :phone, :DOB, :Role)",
                            {
                                fName={cfsqltype:"cf_sql_nvarchar", value:FirstName},
                                lName={cfsqltype:"cf_sql_nvarchar", value:LastName},
                                address={cfsqltype:"cf_sql_nvarchar", value:Address},
                                email={cfsqltype:"cf_sql_nvarchar", value:Email},
                                phone={cfsqltype:"cf_sql_nvarchar", value:Phone},
                                DOB={cfsqltype:"cf_sql_date", value:DOB},
                                Role={cfsqltype:"cf_sql_nvarchar", value:Role}
                            }, {result="addUserResult"});                        
                        }

                        userDetail.Result = arrayToList(result, ", ");
                    }
                    else{
                        var statusErrorAdd = arrayAppend(result, "Email already existed");                        
                        userDetail.Result = arrayToList(result, ", ");
                    }
                }
                else{
                    var statusErrorAdd = arrayAppend(result, "Email id required");
                    userDetail.Result = arrayToList(result, ", ");
                }    
            }            
            addArray = arrayAppend(data, userDetail);
        }

        cffile( action="delete", file=filePath );

        data = arrayToQuery(data);

        data.sort(function(obj1, obj2){
            return compare(obj1.Result, obj2.Result);
        })

        var mySheet = SpreadsheetNew("UserDetails",true);
        SpreadSheetAddRow(mySheet,"First Name,Last Name,Address,Email,Phone,DOB,Role,Result");
        SpreadsheetFormatRow(mySheet, {'bold' : 'true'}, 1);
        
        var i = 1
        cfoutput( query="data" ){
            var FirstName = data["FirstName"];
            var LastName = data["LastName"];
            var Address = data["Address"];
            var Email = data["Email"];
            var Phone = data["Phone"];
            var DOB = data["DOB"];
            var Role = data["Role"];
            var Result = data["Result"];
            if(FirstName != "" OR LastName != "" OR Address != "" OR Email != "" OR Phone != "" OR DOB != "" OR Role != ""){
                spreadsheetSetCellValue(mySheet, FirstName, i+1, 1)
                spreadsheetSetCellValue(mySheet, LastName, i+1, 2)
                spreadsheetSetCellValue(mySheet, Address, i+1, 3)
                spreadsheetSetCellValue(mySheet, Email, i+1, 4)
                spreadsheetSetCellValue(mySheet, Phone, i+1, 5)
                spreadsheetSetCellValue(mySheet, DOB, i+1, 6)
                spreadsheetSetCellValue(mySheet, Role, i+1, 7)
                spreadsheetSetCellValue(mySheet, Result, i+1, 8)
                i++
            }
        }

        cfheader( name="Content-Disposition", value="inline;filename=Upload_Result.xlsx" );
        cfcontent( variable=SpreadSheetReadBinary(mySheet), type="application/vnd.ms-excel" );
    }
}