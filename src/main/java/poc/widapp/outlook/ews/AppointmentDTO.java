package poc.widapp.outlook.ews;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;

import java.util.Date;
import java.util.List;

@Data
public class AppointmentDTO {

    private List<String> attendees;
    @JsonFormat(shape = JsonFormat.Shape.STRING, pattern = "yyyy-MM-dd HH:mm:ss", timezone = "ECT")
    private Date date;
    private int durationInMinutes;
    private String status;

}
