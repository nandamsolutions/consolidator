var App = function() {
    return {
    	totals: {},
    	
        submitForm: function() {
            var formData = new FormData($("form#data")[0]);
            var app = this;
            $.ajax({
                url: "app",
                type: 'POST',
                data: formData,
                async: false,
                success: app.onFilesUpload,
                cache: false,
                contentType: false,
                processData: false
            });
            return false;
        },
        
        onFilesUpload: function(response) {
            $('#form-div').hide();
            $('#view-div').show();
            response = JSON.parse(response);
            var html = _.reduce(_.keys(response), function(html, file) {return html+"<th>"+file+"</th>"}, "<th>&nbsp</th>");
            $('#view-div table thead').html("<tr>"+html+"<th>Total Projected</th></tr>");
            
            var keys = _.reduce(response, function(keys, data) {return _.union(keys, _.keys(data))}, []);
            var tbodyHtml = "";
            _.each(keys, function(key) {
                var total = 0;
                tbodyHtml += "<tr>"+_.reduce(response, function(html, data) {
                    total += data[key];
                    return html +"<td>"+data[key]+"</td>";
                }, "<td>"+key+"</td>")+"<td>"+total+"</td></tr>";
                app.totals[key] = total;
            });
            $('#view-div table tbody').html(tbodyHtml);
        },
        
        downloadConsolidatedReport: function() {
            $('#download').attr('href','/app?data='+JSON.stringify(this.totals));
        }
    };
}