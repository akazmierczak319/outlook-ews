package poc.widapp.outlook.ews;

import lombok.RequiredArgsConstructor;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.LegacyFreeBusyStatus;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.service.ConflictResolutionMode;
import microsoft.exchange.webservices.data.core.enumeration.service.SendInvitationsOrCancellationsMode;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.folder.CalendarFolder;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.CalendarView;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import org.springframework.stereotype.Service;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

@RequiredArgsConstructor
@Service
public class UpdateAppointmentAction {

    private final ExchangeService service;

    public void updateAppointment(String start, String end, String currentSubject, AppointmentDTO updateDTO) throws Exception {
        List<Appointment> appointments = findAppointmentsBetween(start, end);

        // filter by subject name
        Appointment appointmentFound = appointments.stream().filter(appointment -> {
            try {
                return appointment.getSubject().equals(currentSubject);
            } catch (ServiceLocalException e) {
                e.printStackTrace();
            }
            return false;
        }).findFirst().orElseThrow();

        appointmentFound.setSubject(updateDTO.getSubject());
        setNewDate(updateDTO, appointmentFound);
        setStatus(updateDTO, appointmentFound);

        // different options here regarding how to solve eventual conflicts and sending emails
        appointmentFound.update(
                ConflictResolutionMode.AlwaysOverwrite,
                SendInvitationsOrCancellationsMode.SendOnlyToAll
        );

    }

    private void setNewDate(AppointmentDTO updateDTO, Appointment appointmentFound) throws Exception {
        Date newDate = updateDTO.getDate();
        appointmentFound.setStart(newDate);
        Calendar c = Calendar.getInstance();
        c.setTime(newDate);
        c.add(Calendar.MINUTE, updateDTO.getDurationInMinutes());
        appointmentFound.setEnd(c.getTime());
    }

    private List<Appointment> findAppointmentsBetween(String stringStartDate, String stringEndDate) throws Exception {
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date startDate = formatter.parse(stringStartDate);
        Date endDate = formatter.parse(stringEndDate);

        CalendarFolder cf=CalendarFolder.bind(service, WellKnownFolderName.Calendar);
        FindItemsResults<Appointment> findResults =
                cf.findAppointments(new CalendarView(startDate, endDate));

        List<Appointment> result = new ArrayList<>();
        for (Appointment appt : findResults.getItems()) {
            appt.load(PropertySet.FirstClassProperties);
            result.add(appt);
        }

        return result;
    }

    private void setStatus(AppointmentDTO appointmentDTO, Appointment appointment) throws Exception {
        String status = appointmentDTO.getStatus();

        switch (status.toLowerCase()){
            case "busy":
                appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Busy);
                break;
            case "tentative":
                appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Tentative);
                break;
            case "out of office":
                appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.OOF);
                break;
            case "free":
                appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.Free);
                break;
            case "working elsewhere":
                appointment.setLegacyFreeBusyStatus(LegacyFreeBusyStatus.valueOf("Working Elsewhere"));
                break;
        }
    }


}
