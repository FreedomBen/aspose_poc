package com.instructure.aspose_poc;

import com.aspose.words.Document;
import com.aspose.words.Document;

public class App {
    public static void convertWord(String filename) {

    }

    public static void convertPowerpoint(String filename) {
        // Instantiate a Presentation object that represents a presentation file
        Presentation pres = new Presentation(filename);

        // Save the presentation to PDF with default options
        pres.save("demoDefault.pdf", SaveFormat.Pdf);
    }

    public static void convertExcel(String filename) {
	// The path to the documents directory.
	String dataDir = Utils.getSharedDataDir(SaveInPdfFormat.class) + "loading_saving/";

	// Creating an Workbook object with an Excel file path
	Workbook workbook = new Workbook();

	// Save in PDF format
	workbook.save(dataDir + "SIPdfFormat_out.pdf", FileFormatType.PDF);

	// Print Message
	System.out.println("Worksheets are saved successfully.");
    }

    public static void main(String[] args) throws Exception {
        long startTime = System.nanoTime();

        Document doc = null;

        try {
            // Load the document from disk.
            doc = new Document(args[0]);
        } catch(ArrayIndexOutOfBoundsException e) {
            System.out.println("Oops, gotta pass in an input file!");
            System.exit(1);
        }

        // Save the document in PDF format.
        String outputPath = args[0] + ".pdf";
        doc.save(outputPath);

        long endTime = System.nanoTime() - startTime;

        System.out.println("\nDocument converted to PDF successfully.  Took " + ((endTime - startTime) / 1000) + " milliseconds.");
    }
}
