package poc.widapp.outlook.ews;

import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.media.Content;
import io.swagger.v3.oas.annotations.media.ExampleObject;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequiredArgsConstructor
@Slf4j
public class EwsController {

    private final AppointmentAction appointmentAction;

    public static final String EXAMPLE = "{\"attendees\": [\"amaa@gft.com\"],\"date\":\"2021-06-10 12:00:00\", \"durationInMinutes\":60, \"status\":\"free\"}";

    @PostMapping("/ews")
    @Operation(summary = "pass emails addresses to set up meeting",
            requestBody = @io.swagger.v3.oas.annotations.parameters.RequestBody
                (content = @Content(
                        examples = @ExampleObject(value = EXAMPLE))))
    public String appointment(@RequestBody AppointmentDTO appointmentDTO){

        try {
            appointmentAction.executeEws(appointmentDTO);
        } catch (Exception e) {
            log.error("Something went wrong !");
            e.printStackTrace();
            return "Failed";
        }

        return "Success";

    }

}
