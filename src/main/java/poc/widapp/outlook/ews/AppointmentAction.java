package poc.widapp.outlook.ews;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.property.LegacyFreeBusyStatus;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.service.item.Appointment;
import microsoft.exchange.webservices.data.property.complex.AttendeeCollection;
import microsoft.exchange.webservices.data.property.complex.MessageBody;
import org.springframework.stereotype.Service;

import java.util.Calendar;
import java.util.Date;

@RequiredArgsConstructor
@Service
@Slf4j
public class AppointmentAction {

    private final ExchangeService service;

    public void executeEws(AppointmentDTO appointmentDTO) throws Exception {
        Appointment appointment = new Appointment(service);

        appointment.setSubject("Test EWS");
        appointment.setBody(MessageBody.getMessageBodyFromText("test appointment ews !"));

        setDuration(appointmentDTO, appointment);
        setAttendees(appointmentDTO, appointment);
        setStatus(appointmentDTO, appointment);

        appointment.save();

        log.info("appointment saved !");
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

    private void setDuration(AppointmentDTO appointmentDTO, Appointment appointment) throws Exception {
        Date start = appointmentDTO.getDate();
        Date end = appointmentDTO.getDate();
        Calendar c = Calendar.getInstance();
        c.setTime(end);
        c.add(Calendar.MINUTE, appointmentDTO.getDurationInMinutes());
        end = c.getTime();

        appointment.setStart(start);
        appointment.setEnd(end);
    }

    private void setAttendees(AppointmentDTO appointmentDTO, Appointment appointment) throws ServiceLocalException {
        AttendeeCollection requiredAttendees = appointment.getRequiredAttendees();
        appointmentDTO.getAttendees().forEach(email -> {
            try {
                requiredAttendees.add(email);
            } catch (Exception e) {
                e.printStackTrace();
            }
        });
    }

}
