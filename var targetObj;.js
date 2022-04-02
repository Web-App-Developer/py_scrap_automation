var targetObj;
$(".ui-datatable-header").each(function(){
    if($(this).text() === "Published International Application") {
        targetObj = $(this).next().find("a.ps-downloadables").eq(1);
        return;
    }
})