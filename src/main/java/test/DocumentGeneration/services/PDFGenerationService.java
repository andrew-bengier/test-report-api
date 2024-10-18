package test.DocumentGeneration.services;

import com.itextpdf.text.DocumentException;
import org.apache.commons.lang3.*;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Component;
import org.thymeleaf.context.Context;
import org.thymeleaf.spring5.SpringTemplateEngine;
import org.xhtmlrenderer.pdf.ITextRenderer;
import test.DocumentGeneration.types.TemplateCompletionResult;

import java.io.*;
import java.net.*;
import java.text.*;
import java.time.*;
import java.time.format.*;
import java.util.*;

@Component
public class PDFGenerationService {
    //static SpringTemplateEngine templateEngine = new SpringTemplateEngine();
    final DateTimeFormatter INSTANT_FORMATTER = DateTimeFormatter.ISO_LOCAL_DATE.withZone(ZoneOffset.UTC);
    final SimpleDateFormat DATE_FORMATTER = new SimpleDateFormat("MM-dd-yyyy");

    private final SpringTemplateEngine templateEngine;

    public PDFGenerationService(SpringTemplateEngine templateEngine) {
        this.templateEngine = templateEngine;
    }

    public TemplateCompletionResult completePDFGeneration() {
        // Content for Template
        Context context = PDFDocumentUtil.setupData();
        // Template
        String templateBody = templateEngine.process("/templates/tableDocument", context);

        // Prepare pdf renderer
        ITextRenderer renderer = new ITextRenderer();
        renderer.setDocumentFromString(templateBody);
        renderer.layout();

        // Generate the PDF and attach it to the response
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            renderer.createPDF(outputStream);
        } catch (DocumentException e) {
            System.out.println("Error while writing in the Account Statement template");
        } catch (IOException e) {
            System.out.println("Error while accessing the Account Statement template");
        }

        return TemplateCompletionResult.builder()
                .completedDocument(outputStream)
                .build();
    }

    public byte[] testPoi() throws InvalidFormatException, IOException, URISyntaxException {
        // Fill test data
        Map<String, Object> data = new HashMap<>();
        data.put("today", INSTANT_FORMATTER.format(Instant.now()));
        data.put("lob", "General Casualty, Premises Ops");
        data.put("claimNumber", "RN-9-0000883");
        data.put("policyNumber", "RN-7-0300083");
        data.put("examiner", "Andrew Bengier");
        data.put("policyPeriod", formatTimePeriod(new Date("10/01/2023"), new Date("10/01/2024")));
//        <today>
//        <applicablePolLimit>
//        <lob>
//        <claimNumber>
//        <aggregate>
//        <policyNumber>
//        <examiner>
//        <claimExpense>
//        <policyPeriod>
//        <underwriter>
//        <broker>
//        <dateRetro>
//        <deductible>
//        <stateInsured>
//        <stateLoss>
//        <insured>
//        <claimants>
//        <dateLoss>
//        <dateClaim>
//        <causeOfLoss>
//        <dateRNNotified>
//        <catNumber>
//        <litigation>
//        <dateMediation>
//        <dateTrial>
//
//        <amtReserveLoss>
//        <amtPaidLoss>
//        <amtReserveLoss>
//        <amtIncurredLoss>
//        <amtReserveExpense>
//        <amtPaidExpense>
//        <amtExpenseLoss>
//        <amtIncurredExpense>
//        <amtReserveTotal>
//        <amtPaidTotal>
//        <amtTotalLoss>
//        <amtIncurredTotal>
//
//        <amtRnReserveLoss>
//        <amtRnPaidLoss>
//        <amtRnReserveLoss>
//        <amtRnIncurredLoss>
//        <amtRnReserveExpense>
//        <amtRnPaidExpense>
//        <amtRnExpenseLoss>
//        <amtRnIncurredExpense>
//        <amtRnReserveTotal>
//        <amtRnPaidTotal>
//        <amtRnTotalLoss>
//        <amtRnIncurredTotal>
//        <loss>
//        <coverage>
//        <liability>
//        <damages>
//        <reserves>
//        <resolution>


        URL resource = getClass().getClassLoader().getResource("templates/RN_claims_reinsurance_report.docx");
        if (resource == null) {
            throw new IllegalArgumentException("file not found!");
        } else {

            // failed if files have whitespaces or special characters
            //return new File(resource.getFile());

            File template = new File(resource.toURI());
            XWPFDocument doc = new XWPFDocument(OPCPackage.open(template));

            for (XWPFTable tbl : doc.getTables()) {
                for (XWPFTableRow row : tbl.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun r : p.getRuns()) {
                                String text = r.getText(0);
                                if (text != null) {
                                    for(String tag : data.keySet()){
                                        if(text.contains(tag)){
                                            text = text.replace(tag, data.get(tag).toString());
                                            r.setText(text, 0);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            doc.write(outputStream);
            doc.close();

            return outputStream.toByteArray();
        }
    }

    private String formatTimePeriod(Date effective, Date expiration) {
        if (ObjectUtils.isNotEmpty(effective) && ObjectUtils.isNotEmpty(expiration)) {
            return DATE_FORMATTER.format(effective) + " to " + DATE_FORMATTER.format(expiration);
        } else if (ObjectUtils.isNotEmpty(effective)) {
            return "Effective " + DATE_FORMATTER.format(effective);
        } else {
            return "Expires " + DATE_FORMATTER.format(expiration);
        }
    }
}
