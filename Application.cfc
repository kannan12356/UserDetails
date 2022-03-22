component{

    this.name = "User details";
    this.datasource = "UserDetails";
    this.sessionManagement  = true;
    this.sessionTimeout = CreateTimeSpan(0, 0, 30, 0);
    this.mappings["/local"] = getDirectoryFromPath(getCurrentTemplatePath());

}