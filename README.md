## Certificate generation and emailing project

---

### Purpose

This project is about building a workflow in UiPath studio to automate the repetitive task of generating participation certificates for participants participating in an event.

---

### Required Packages

1. BalaReva.EasyPowerPoint.Activities
2. BalaReva.PowerPoint.Activities

---

### Scenario

1. We have details of participants (name and email ID) stored in an excel sheet.
2. A certificate template is available in the form of a PowerPoint slide.
3. The excel sheet is read using the Read Range activity and stored in a variable of type DataTable.
4. Loop through the DataTable and extract the name of each participant.
5. Use the name and write it to the certificate template using the TextShapeEdit activity.
6. Export the slide as PDF using the Export To PDF activity and store it in a folder where all the certificates are supposed to be stored.
7. Loop through the DataTable and extract the name and email of the participant.
8. Find the generated certificate of the participant and send it as an attachment to the participants email ID using the Send SMTP mail message activity.

---

### Workflow

![Workflow](https://i.ibb.co/NL1WYyD/Workflow.png)
