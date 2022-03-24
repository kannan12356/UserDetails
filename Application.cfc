component{

    this.name = "User details";
    this.datasource = "UserDetails";
    this.sessionManagement  = true;
    this.sessionTimeout = CreateTimeSpan(0, 0, 30, 0);
    this.mappings["/local"] = getDirectoryFromPath(getCurrentTemplatePath());

    function onError(Exception,EventName){
        writeOutput('<center><h1>An error occurred</h1>
        <p>Please Contact the developer</p>
        <p>Error details: #Exception.message#</p></center>');
    } 

}