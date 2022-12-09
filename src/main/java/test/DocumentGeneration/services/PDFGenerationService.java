package test.DocumentGeneration.services;

import com.itextpdf.text.DocumentException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.thymeleaf.TemplateEngine;
import org.thymeleaf.context.Context;
import org.thymeleaf.spring5.SpringTemplateEngine;
import org.thymeleaf.templateresolver.ClassLoaderTemplateResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;
import test.DocumentGeneration.types.TemplateCompletionResult;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@Component
public class PDFGenerationService {
    static SpringTemplateEngine templateEngine = new SpringTemplateEngine();

    public static TemplateCompletionResult completePDFGeneration() {
        ClassLoaderTemplateResolver templateResolver= new ClassLoaderTemplateResolver();
        templateResolver.setSuffix(".html");
        templateResolver.setTemplateMode("HTML");
        templateResolver.setCacheable(false);
        templateEngine.setTemplateResolver(templateResolver);
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
}
