package poc.widapp.outlook.ews;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.ExampleObject;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequiredArgsConstructor
@Slf4j
public class EwsController {

    private final AppointmentAction appointmentAction;
    private final UpdateAppointmentAction updateAppointmentAction;

    public static final String EXAMPLE = "{\"subject\":\"test ews\",\"attendees\": [\"amaa@gft.com\"],\"date\":\"2021-06-10 12:00:00\", \"durationInMinutes\":60, \"status\":\"free\"}";
    public static final String EXAMPLE_UPDATE = "{\n" +
            "  \"date\": \"2021-06-16 12:00:00\",\n" +
            "  \"durationInMinutes\": 60,\n" +
            "  \"status\": \"busy\",\n" +
            "  \"subject\": \"updated\"\n" +
            "}";

    @PostMapping("/ews")
    @Operation(summary = "pass emails addresses to set up meeting",
            requestBody = @io.swagger.v3.oas.annotations.parameters.RequestBody
                (content = @Content(
                        examples = @ExampleObject(value = EXAMPLE))))
    public String appointment(@RequestBody AppointmentDTO appointmentDTO){

        try {
            return appointmentAction.executeEws(appointmentDTO);
        } catch (Exception e) {
            log.error("Something went wrong !");
            e.printStackTrace();
            return "Failed";
        }
        
    }

    @PostMapping("/ews/update/{start}/start/{end}/end/{subject}/subject")
    @Operation(summary = "pass parameters to update appointment with given subject " +
            "occurring between passed dates",
            requestBody = @io.swagger.v3.oas.annotations.parameters.RequestBody
                    (content = @Content(
                            examples = @ExampleObject(value = EXAMPLE_UPDATE))))
    public String update(@PathVariable String start, @PathVariable String end,
                              @PathVariable String subject, @RequestBody AppointmentDTO updateDTO){

        try {
            updateAppointmentAction.updateAppointment(start, end, subject, updateDTO);
        } catch (Exception e) {
            log.error("Something went wrong !");
            e.printStackTrace();
            return "Failed";
        }

        return "Success";
    }
    
}
