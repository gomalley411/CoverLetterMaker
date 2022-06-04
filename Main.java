package com.gomalley411.wordprojectmaker;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Scanner;

// Java program to create a Word document
// Importing Spire Word libraries
 
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;
 
class Main {
 
    // Main driver method
    public static void main(String[] args) {

        // stores date as a string
        DateFormat df = new SimpleDateFormat("MM-dd-yyyy");
        Calendar cal = Calendar.getInstance(); 
        Date date = cal.getTime();
        String todaysDate = df.format(date);

        // create a Word document
        Document document = new Document();

        // create body style and address style
        ParagraphStyle body = new ParagraphStyle(document);
        body.setName("mybody");
        body.getCharacterFormat().setFontName("Calibri");
        body.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(body);

        ParagraphStyle addr = new ParagraphStyle(document);
        addr.setName("myaddr");
        addr.getCharacterFormat().setFontName("Calibri");
        addr.getCharacterFormat().setFontSize(11f);
        document.getStyles().add(addr);
 
        // Add a section
        Section section = document.addSection();

        // set margins
        section.getPageSetup().getMargins().setAll(65f);

        // add address, date and info. These don't change very often so be sure to fill them in before running code
        Paragraph a = section.addParagraph();
        a.appendText("{{YOUR NAME HERE}}\n");
        a.appendText("{{STREET}}\n{{CITY}}, {{STATE}} {{ZIP}}\n\n");
        a.appendText(todaysDate + "\n");

        a.applyStyle("myaddr");

        Scanner kb = new Scanner(System.in);
        System.out.println("Enter the position name:");
        String posName = kb.nextLine();
        System.out.println("Enter the company name:");
        String companyName = kb.nextLine();

        // greeting
        Paragraph greeting = section.addParagraph();
        greeting.appendText("To the hiring manager or HR team at " + companyName + ",");
        greeting.applyStyle("mybody");

        // intro
        Paragraph p1 = section.addParagraph();
        p1.appendText("I am writing to you today regarding the " + posName + " position at " + companyName + ". I saw the job posting on your website and immediately thought it would be the perfect opportunity for me. I would be greatly honored to be considered for this position at your company.");
        p1.applyStyle("mybody");

        // paragraph 2
        Paragraph p2 = section.addParagraph();
        p2.appendText("{{QUALIFICATIONS AND SKILLS HERE}}");
        p2.applyStyle("mybody");

        // ask if user would like to add a third paragraph if needed
        System.out.println("Would you like to add a third paragraph? Answer 1 for yes or 2 for no.");
        int choice = kb.nextInt();
        if (choice == 1) {
            System.out.println("Please enter what you would like to say.");
            kb.nextLine();
            String thirdParagraph = kb.nextLine();
            Paragraph p3 = section.addParagraph();
            p3.appendText(thirdParagraph);
            p3.applyStyle("mybody");
        }
        kb.close();

        // conclusion
        Paragraph conclusion = section.addParagraph();
        conclusion.appendText("I have included my resume with this application. I am available for an interview at your convenience. You can contact me via email at {{EMAIL HERE}} or by phone at the number listed above. Thank you for your time and consideration.");
        conclusion.applyStyle("mybody");

        // signing
        Paragraph signing = section.addParagraph();
        signing.appendText("Sincerely,\n{{FULL NAME HERE}}");
        signing.applyStyle("myaddr");

        // Iteration for white spaces
        for (int i = 0;
             i < section.getParagraphs().getCount(); i++) {
                 if (section.getParagraphs().get(i).getStyle().getName().equals("mybody")) {
                     section.getParagraphs().get(i).getFormat().setAfterAutoSpacing(true);
                 }
        }

        // change line 112 to wherever you want your document to be stored on your computer
        String saveAddr = "com\\gomalley411\\wordprojectmaker\\cover letters\\" 
        + companyName.toLowerCase() + " cover letter " + todaysDate + ".docx";

        // Save the document
        document.saveToFile(saveAddr,FileFormat.Docx);
        System.out.println("Saved to " + saveAddr);
    }
}
