<cfinvoke
    component="cfc/userDetails"
    method="getData"
    returnVariable="userDetails">
</cfinvoke>

<cfif structKeyExists(form, "upload")>
    <cfinvoke 
        component="cfc/userDetails"
        method="addData"
        returnVariable="addResult">
        <cfinvokeargument name="file" value="#form.file#">
    </cfinvoke>
    <cflocation  url="index.cfm" addToken="false">
</cfif>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="assets/bootstrap/css/bootstrap.min.css"/>
    <link rel="stylesheet" href="assets/style.css"/>
    <title>User Details</title>
</head>
<body>
    <main>
        <h4>USER INFORMATION</h4>
        <div class="buttons">
            <div class="dwnld-buttons">
                <a href="cfc/userDetails.cfc?method=downloadPlainTemplate"><button class="btn btn-info btn-sm">Plain Template</button></a>
                <a href="cfc/userDetails.cfc?method=downloadData"><button class="btn btn-primary btn-sm">Template with data</button></a>
            </div>
            <form action="" method="post" enctype="multipart/form-data">
                <input type="file" id="file" name="file"  accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel" required/>
                <input class="btn btn-secondary btn-sm" type="button" value="Browse" onclick="document.getElementById('file').click();" />
                <button type="submit" name="upload" class="btn btn-success btn-sm">Upload</button>
            </form>
        </div>

        <div class="user-detail">
            <table class="table table-bordered">
                <thead>
                    <th>First Name</th>
                    <th>Last Name</th>
                    <th>Address</th>
                    <th>Email</th>
                    <th>Phone</th>
                    <th>DOB</th>
                    <th>Role</th>
                </thead>         
                <tbody>
                    <cfoutput query="userDetails">
                        <tr>
                            <td>#FirstName#</td>
                            <td>#LastName#</td>
                            <td>#Address#</td>
                            <td>#Email#</td>
                            <td>#Phone#</td>
                            <td>#dateFormat(DOB, "dd-mm-YYYY")#</td>
                            <td>#Role#</td>
                        </tr>
                    </cfoutput>
                </tbody>   
            </table>
        </div>
    </main>
</body>
</html>