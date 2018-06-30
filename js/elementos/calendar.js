$(document).ready(function () {

    //var url_string = window.location.href;
    //var url = new URL(url_string);
    var user = getParameterByName("user");
    var name = getParameterByName("name");

    if (user != "0" && user != null) {
        $('#titulo').text("Calendario de Actividades de " + name);
        $('#cssmenu').hide();
    }
       

    combobox($('#cbxIngeniero'), { 1: 'inspectores' }, 'Seleccione Ingeniero...');
    setTimeout(function () { $('#cbxIngeniero').val(user) }, 1000)
    $('#cbxIngeniero').change(function () {
        LoadCalendar($('#cbxIngeniero').val(), true);
    });

    LoadCalendar(user || '0', false);

});

function LoadCalendar(ing, refresh) {
    var currentdate = new Date();
    var datetime = currentdate.getFullYear() + "-" + (currentdate.getMonth() + 1) + "-" + (currentdate.getDay());

    if (refresh) {

        $.ajax({
            url: 'handlers/calendar.ashx',
            type: 'POST',
            data: { 1: 'getCalendarData', ingeniero: ing },
            success: function (result) {
                if (result.done) {
                    console.log(result.data)
                    $('#calendar').fullCalendar('removeEvents');
                    $('#calendar').fullCalendar('addEventSource', result.data);
                    $('#calendar').fullCalendar('rerenderEvents');
                }
                else {
                    alertify.error(result.message);
                }

            }
        });

    }
    else {
        $.ajax({
            url: 'handlers/calendar.ashx',
            type: 'POST',
            data: { 1: 'getCalendarData', ingeniero: ing },
            success: function (result) {
                if (result.done) {
    
                    $('#calendar').fullCalendar({
                        header: {
                            left: 'prev,next today',
                            center: 'title',
                            right: 'month,basicWeek,basicDay'
                        },
                        locale: 'es',
                        //defaultDate: datetime,
                        navLinks: true, // can click day/week names to navigate views
                        editable: true,
                        eventLimit: true, // allow "more" link when too many events
                        timeFormat: 'H(:mm)',
                        events: result.data,
                        eventRender: function (event, element) {
                            $(element).prop({ title: event.ingeniero }).tooltip()
                        }
                    });

                    console.log(result.data)
                }
                else {
                    alertify.error(result.message);
                }

            }
        });
    }

}
