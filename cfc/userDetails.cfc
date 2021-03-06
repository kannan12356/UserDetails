component{

    public function getData(){
        var userDetails = queryExecute("Select Users.UserId, FirstName, LastName, Address, Email, Phone, 
        DOB, roles.RoleId AS RoleId, GROUP_CONCAT(roles.Role) as Role From Users
        INNER JOIN user_roles ON Users.UserId=user_roles.UserId
        INNER JOIN roles ON roles.RoleId=user_roles.RoleId
        GROUP BY users.UserId;",[]);

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
        var dataQuery = arrayReduce(data, function(newQuery, element) {
                            structEach(element, function(key) {
                                if (!newQuery.keyExists(key)) {
                                    queryAddColumn(newQuery, key, []);
                                }
                            });
                            queryAddRow(newQuery, element);
                            return newQuery;
                        }, QueryNew(""));
        return dataQuery;
    }

    private function addUserRole(required Roles,required userId){
        for(i=1; i<=arrayLen(arguments.Roles); i++){
            var insertUserRole = queryExecute("Insert into user_roles
            (UserId, RoleId) VALUES (:userId, (select RoleId from roles WHERE Role=:Role))",
            {
                userId={cfsqltype:"cf_sql_integer", value:arguments.userId},
                Role={cfsqltype:"cf_sql_nvarchar", value:Roles[i]}
            }, {result="userRoleAddResult"});
        }
    }

    private function createExcel(required data){
        data = arrayToQuery(data);
        data.sort(function(obj1, obj2){
            return compare(obj1.Success, obj2.Success);
        });
        var excelSheet = SpreadsheetNew("UserDetails",true);
        SpreadSheetAddRow(excelSheet,"First Name,Last Name,Address,Email,Phone,DOB,Role,Result");
        SpreadsheetFormatRow(excelSheet, {'bold' : 'true'}, 1);

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
                spreadsheetSetCellValue(excelSheet, FirstName, i+1, 1)
                spreadsheetSetCellValue(excelSheet, LastName, i+1, 2)
                spreadsheetSetCellValue(excelSheet, Address, i+1, 3)
                spreadsheetSetCellValue(excelSheet, Email, i+1, 4)
                spreadsheetSetCellValue(excelSheet, Phone, i+1, 5)
                spreadsheetSetCellValue(excelSheet, DOB, i+1, 6)
                spreadsheetSetCellValue(excelSheet, Role, i+1, 7)
                spreadsheetSetCellValue(excelSheet, Result, i+1, 8)
                i++
            }
        }
        return excelSheet;
    }
    
    public function addData(required file){
        cffile( fileField="file", nameconflict="overwrite", destination="E:\ColdFusion\cfusion\wwwroot\UserDetails\Excel", action="upload", result="uploadFile" );
        var filePath = "#uploadFile.serverDirectory#\#uploadFile.clientFile#";

        cfspreadsheet( action="read", src="#filePath#", excludeheaderrow="true", headerrow=1, query="excelData" );

        var data = arrayNew(1);
        var emails = arrayNew(1);
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
                userDetail.Success = 0;
                
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
                                else{
                                    var RoleErrorAdd = arrayAppend(result, "Role not matched in database");
                                }
                            }
                            
                            if(!arrayIsEmpty(Roles)){          
                                var addEmail = arrayAppend(emails, Email);
                                var insertData = queryExecute("INSERT into Users 
                                (FirstName, LastName, Address, Email, Phone, DOB) 
                                Values 
                                (:fName, :lName, :address, :email, :phone, :DOB)",
                                {
                                    fName={cfsqltype:"cf_sql_nvarchar", value:FirstName},
                                    lName={cfsqltype:"cf_sql_nvarchar", value:LastName},
                                    address={cfsqltype:"cf_sql_nvarchar", value:Address},
                                    email={cfsqltype:"cf_sql_nvarchar", value:Email},
                                    phone={cfsqltype:"cf_sql_nvarchar", value:Phone},
                                    DOB={cfsqltype:"cf_sql_date", value:DOB}
                                }, {result="addUserResult"});  
                                var lastInsertId = addUserResult.generatedKey;
                                var addUserRoles = addUserRole(Roles, lastInsertId);                                
                                var statusErrorAdd = arrayAppend(result, "Success");
                                userDetail.success = 1;
                            }                            
                        }
                        userDetail.Result = arrayToList(result, ", ");
                    }
                    else{
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
                            else{
                                var RoleErrorAdd = arrayAppend(result, "Role not matched in database");
                            }
                        }
                        if(!arrayIsEmpty(Roles)){                         
                            if(!arrayContains(emails, Email)){
                                var addEmail = arrayAppend(emails, Email);
                                var updateData = queryExecute("UPDATE Users SET
                                FirstName=:fName, LastName=:lName, Address=:address, Email=:email,
                                Phone=:phone, DOB=:DOB WHERE email=:email",
                                {
                                    fName={cfsqltype:"cf_sql_nvarchar", value:FirstName},
                                    lName={cfsqltype:"cf_sql_nvarchar", value:LastName},
                                    address={cfsqltype:"cf_sql_nvarchar", value:Address},
                                    email={cfsqltype:"cf_sql_nvarchar", value:Email},
                                    phone={cfsqltype:"cf_sql_nvarchar", value:Phone},
                                    DOB={cfsqltype:"cf_sql_date", value:DOB}
                                }, {result="updateResult"});
                                
                                var getUserId =  queryExecute("SELECT UserId FROM Users WHERE email=:email",
                                {email={cfsqltype:"cf_sql_nvarchar", value:email}},{result="getUserIdResult"})
                                var userId = getUserId.UserId;

                                var deleteRoles = queryExecute("DELETE from user_roles WHERE UserId=:userId",
                                {userId={cfsqltype:"cf_sql_integer", value:userId}}, {result="deleteRoleResult"});
                                var addUserRoles = addUserRole(Roles, userId);
                                
                                var statusErrorAdd = arrayAppend(result, "Updated");
                                userDetail.success = 1;
                            }
                            else{
                                var statusErrorAdd = arrayAppend(result, "Email id already existed");
                            }  
                        }
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
        var mySheet = createExcel(data);
        cfheader( name="Content-Disposition", value="inline;filename=Upload_Result.xlsx" );
        cfcontent( variable=SpreadSheetReadBinary(mySheet), type="application/vnd.ms-excel" );
    }
}