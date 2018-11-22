// Example insert text into editors (different implementations)
(function(window, undefined){
    
    var text = "Hello world!";

    window.Asc.plugin.init = function()
    {
        var variant = 1;

        switch (variant)
        {
            case 0:
            {
                var sScript = "var Haha = Api.GetActiveSheet();";
                sScript += "var col = Haha.GetActiveCell().GetCol();";
                sScript += "var row = Haha.GetActiveCell().GetRow();";
                sScript += "var cur = Haha.GetActiveCell();";
                sScript += "cur.SetValue('hello');";
                sScript += "oColor = Api.CreateColorFromRGB(49, 133, 154);";
                sScript += "cur.SetFontColor(oColor);";
                sScript += "cur.SetFillColor(Api.CreateColorFromRGB(0, 0, 250));";
                this.info.recalculate = true;
                this.executeCommand("close", sScript);
                break;
            }
            case 1:
            {
                // call command without external variables
                this.callCommand(function() {
                    var Haha = Api.GetActiveSheet();
                    var col = Haha.GetActiveCell().GetCol();
                    var row = Haha.GetActiveCell().GetRow();
                    var cur = Haha.GetActiveCell();
                    cur.SetValue("hello");
                    oColor = Api.CreateColorFromRGB(49, 133, 154);
                    cur.SetFontColor(oColor);
                    cur.SetFillColor(Api.CreateColorFromRGB(0, 0, 250));
                }, true);
                break;
            }
            default:
                break;
        }
    };

    window.Asc.plugin.button = function(id)
    {
    };

})(window, undefined);
