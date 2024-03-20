package org.example;


import com.amazonaws.auth.AWSCredentialsProvider;
import com.amazonaws.auth.DefaultAWSCredentialsProviderChain;
import com.amazonaws.regions.Regions;
import com.amazonaws.services.comprehend.AmazonComprehend;
import com.amazonaws.services.comprehend.AmazonComprehendClientBuilder;
import com.amazonaws.services.comprehend.model.*;
import com.amazonaws.services.translate.AmazonTranslate;
import com.amazonaws.services.translate.AmazonTranslateClientBuilder;
import com.amazonaws.services.translate.model.TranslateTextRequest;
import com.amazonaws.services.translate.model.TranslateTextResult;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;


import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;



public class Project {
    public static void main(String[] args) {

        AWSCredentialsProvider credentialsProvider = new DefaultAWSCredentialsProviderChain();
        AmazonComprehend comprehendClient = AmazonComprehendClientBuilder.standard()
                .withCredentials(credentialsProvider)
                .withRegion(Regions.US_EAST_1)
                .build();
        AmazonTranslate translateClient = AmazonTranslateClientBuilder.standard()
                .withCredentials(credentialsProvider)
                .withRegion(Regions.US_EAST_1)
                .build();



        IOUtils.setByteArrayMaxOverride(600_000_000);

        String inputFile = "C:\\Users\\pc\\Desktop\\PFE\\Akhbarona.xlsx";

        try (FileInputStream fileInputStream = new FileInputStream(new File(inputFile))) {
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            int iteration = 1;

            for (Row row : sheet) {

                // Create a new XML document
                DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
                DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
                Document doc = docBuilder.newDocument();

                // Create the root element
                Element rootElement = doc.createElement("results");
                doc.appendChild(rootElement);

                String fileName = "output_" + iteration + ".xml";
                String filePath = "C:\\Users\\pc\\Desktop\\Result\\" + fileName;

                Cell body = row.getCell(1);
                String text = body.getStringCellValue();

                // Detect the dominant language
                DetectDominantLanguageRequest detectLanguageRequest = new DetectDominantLanguageRequest()
                        .withText(text);
                DetectDominantLanguageResult detectLanguageResult = comprehendClient.detectDominantLanguage(detectLanguageRequest);

                List<DominantLanguage> dominantLanguages = detectLanguageResult.getLanguages();
                if (!dominantLanguages.isEmpty()) {
                    String detectedLanguageCode = dominantLanguages.get(0).getLanguageCode();

                    if (detectedLanguageCode.equals("ar")) {
                        // Create XML elements for each iteration
                        Element textElement = doc.createElement("Text");
                        Element detectedLanguageElement = doc.createElement("DetectedLanguage");
                        Element translationElement = doc.createElement("TranslatedText");
                        Element TarsentimentElement = doc.createElement("TargetedSentiment");
                        Element syntaxAnalysisElement = doc.createElement("SyntaxAnalysis");

                        int maxTextSize = 2000;
                        int delay = 2000; // Delay between each translation request in milliseconds
                        StringBuilder translationBuilder = new StringBuilder();
                        StringBuilder sentimentBuilder = new StringBuilder();
                        StringBuilder targetedsentimentBuilder = new StringBuilder();
                        List<String> syntaxTokens = new ArrayList<>();

                        for (int i = 0; i < text.length(); i += maxTextSize) {
                            int endIndex = Math.min(i + maxTextSize, text.length());
                            String part = text.substring(i, endIndex);

                            textElement.setTextContent(part);

                            // Translate the text to a desired language
                            String targetLanguageCode = "en";
                            TranslateTextRequest translateRequest = new TranslateTextRequest()
                                    .withText(part)
                                    .withSourceLanguageCode(detectedLanguageCode)
                                    .withTargetLanguageCode(targetLanguageCode);
                            TranslateTextResult translateResult = translateClient.translateText(translateRequest);

                            String translatedText = translateResult.getTranslatedText();
                            translationBuilder.append(translatedText).append("\n");


                            //Targeted sentiment
                            DetectTargetedSentimentRequest TardetectSentimentRequest = new DetectTargetedSentimentRequest()
                                    .withLanguageCode(targetLanguageCode)
                                    .withText(translatedText);
                            DetectTargetedSentimentResult TardetectSentimentResult = comprehendClient.detectTargetedSentiment(TardetectSentimentRequest);
                            targetedsentimentBuilder.append(TardetectSentimentResult).append("\n");

                            // Syntax analysis
                            DetectSyntaxRequest detectSyntaxRequest = new DetectSyntaxRequest()
                                    .withLanguageCode(targetLanguageCode)
                                    .withText(translatedText);
                            DetectSyntaxResult detectSyntaxResult = comprehendClient.detectSyntax(detectSyntaxRequest);
                            List<SyntaxToken> tokens = detectSyntaxResult.getSyntaxTokens();
                            for (SyntaxToken token : tokens) {
                                syntaxTokens.add(token.getText() + " (" + token.getPartOfSpeech().getTag() + ")");
                            }

                            try {
                                Thread.sleep(delay);
                            } catch (InterruptedException e) {
                                e.printStackTrace();
                            }
                        }

                        // Set the text content of the XML elements
                        detectedLanguageElement.setTextContent(detectedLanguageCode);
                        translationElement.setTextContent(translationBuilder.toString());
                        TarsentimentElement.setTextContent(targetedsentimentBuilder.toString());
                        syntaxAnalysisElement.setTextContent(String.join(", ", syntaxTokens));

                        // Append the XML elements to the root element
                        rootElement.appendChild(textElement);
                        rootElement.appendChild(detectedLanguageElement);
                        rootElement.appendChild(translationElement);
                        rootElement.appendChild(TarsentimentElement);
                        rootElement.appendChild(syntaxAnalysisElement);

                        // Write the XML document to the output file


                        TransformerFactory transformerFactory = TransformerFactory.newInstance();
                        Transformer transformer = transformerFactory.newTransformer();
                        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
                        DOMSource source = new DOMSource(doc);
                        StreamResult result = new StreamResult(new File(filePath));
                        transformer.transform(source, result);
                        iteration++;
                    } else {
                        System.out.println("The text is not in Arabic");
                    }
                }
            }

            // Shutdown AWS clients to release resources
            comprehendClient.shutdown();
            translateClient.shutdown();
        } catch (IOException | ParserConfigurationException | TransformerException e) {
            e.printStackTrace();
        }
    }
}
