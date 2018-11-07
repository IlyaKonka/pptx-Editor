package parser;

import com.codahale.metrics.Timer;
import com.twelvemonkeys.image.AffineTransformOp;
import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.util.Zip4jConstants;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xslf.usermodel.*;
import parser.Mail;

import javax.imageio.ImageIO;
import javax.mail.MessagingException;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Change .pptx presentation
 */

public class PptxParser {

    /**
     * This method backups the presentations.
     *
     * @param  zipFolder  the path to zip file with all pictures from presentations
     * @param  zipPassword zip archive password
     * @param  presentationListFolder the path to list of the presentations which should be backup
     * @note   the list of presentations should contain all names of presentations
     */


    public static void backup(String zipFolder, String zipPassword, String presentationListFolder) {
        ArrayList<String> picNamesArr = new ArrayList<>();
        String dest = null;
        try {
            ZipFile zipFile = new ZipFile(zipFolder);
            if (zipFile.isEncrypted()) {
                zipFile.setPassword(zipPassword);
            }
            dest = zipFolder.substring(0, zipFolder.lastIndexOf("\\")) + "\\pic";
            zipFile.extractAll(dest);
        } catch (ZipException e) {
            System.out.println("Back up zip file with the pictures can not be found.");
        }


        ArrayList<String> presentationsToBackUp = null;
        try {
            presentationsToBackUp = (ArrayList<String>)
                    Files.readAllLines(Paths.get(presentationListFolder));
        } catch (IOException e) {
            System.out.println("Presentation list can not be found.");
        }


        boolean isFound = false;
        int counterBackedUpPresentation = 0;
        File picDestFolder = new File(dest);
        ArrayList<File> picList = new ArrayList<File>(Arrays.asList(picDestFolder.listFiles()));
        Comparator comparator = new Comparator<File>() {
            @Override
            public int compare(File o1, File o2) {
                return o1.getName().compareTo(o2.getName());
            }
        };
        picList.sort(comparator);

        ArrayList<File> picListCopy = new ArrayList<>(picList);
        for (String s : presentationsToBackUp) {
            isFound = false;
            XMLSlideShow ppt;
            try {
                ppt = new XMLSlideShow(new FileInputStream(s));
            } catch (Exception exc) {
                System.out.print(exc + " ||| ");
                System.out.println("Can't find a presentation: " + s);
                System.out.println("Back up crashed");
                return;
            }


            //TODO from here
            // get slides
            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape sh : slide.getShapes()) {
                    if (sh instanceof XSLFPictureShape) {
                        XSLFPictureShape shape = (XSLFPictureShape) sh;
                        XSLFPictureData pictureData = shape.getPictureData();
                        int picHeight = pictureData.getImageDimensionInPixels().height;
                        int picWidth = pictureData.getImageDimensionInPixels().width;
                        String fileName = pictureData.getFileName();
                        if ((picHeight > 100 && picWidth > 100) && !picNamesArr.contains(fileName)) {
                            picNamesArr.add(fileName);
                            try {
                                byte[] bytes = pictureData.getData();
                                BufferedImage img = ImageIO.read(new ByteArrayInputStream(bytes));
                                if (img == null)
                                    throw new Exception();
                                try {
                                    byte[] picContent = Files.readAllBytes(picList.get(0).toPath());
                                    pictureData.setData(picContent);
                                    isFound = true;
                                    try {
                                        picList.remove(0);
                                    } catch (Exception e) {
                                        System.out.print(e + " ||| ");
                                        System.out.println("No more pictures to back up.");
                                        return;
                                    }
                                } catch (Exception e) {
                                    System.out.print(e + " ||| ");
                                    System.out.println("Problem with back up. Try again.");
                                    return;
                                }
                            } catch (Exception e) {
                                continue;
                            }
                        }
                    }
                }
            }

            //save changes
            try {
                if (isFound) {
                    counterBackedUpPresentation++;
                    FileOutputStream out = new FileOutputStream(s);
                    ppt.write(out);
                    System.out.println("Presentation \"" + s.substring(s.lastIndexOf("\\") + 1) + "\" backed up successfully");
                    out.close();
                }
            } catch (Exception saveExc) {
                System.out.print(saveExc + " ||| ");
                System.out.println("Can't save a backed up file: " + s);
            }
        }

        //delete pic from pic
        for (File f : picListCopy) {
            try {
                Files.delete(f.toPath());
            } catch (IOException e) {
                System.out.print(e + " ||| ");
                System.out.println("Can not be deleted: " + f.getAbsolutePath());
            }
        }
        //delete pic folder
        try {
            Files.delete(Paths.get(dest));
        } catch (IOException e) {
            System.out.print(e + " ||| ");
            System.out.println("Can not be deleted: " + dest);
        }

        System.out.println("Presentations backed up: " + counterBackedUpPresentation);
    }

    /**
     * This method main method. It is used to change the presentations.
     *
     * @param  presentationPath  the path to folder with presentations of one pptx presentation
     * @param  oldWord the word which should be changed
     * @param  newWord the new word
     * @param  pictureMode true if picture mode on
     * @param  pictureText if pictureMode is ON, the text on the pictures. else null
     * @param  style if pictureMode is ON, the text style. else null
     * @param  color if pictureMode is ON, the color of text. else null
     * @param  manyTimesWrite if pictureMode is ON and manyTimesWrite is true,
     *                       the text will be many time written on the pictures. else does't matter
     * @param  toRotate if pictureMode is ON, all fotos will be 180 deg rotated
     * @param  toSave true for backup
     * @param  zipFolder  the path to zip file with all pictures from presentations
     * @param  zipPassword zip archive password
     * @param  presentationListFolder the path to list of the presentations which should be backup
     * @param  receiverMailAdr mail address to send info about backup. if null, no info will be sent.
     * @note   the list of presentations should contain all names of presentations
     */

    public static void parse(String presentationPath, String oldWord, String newWord, boolean pictureMode,
                             String pictureText, String style, String color,
                             boolean manyTimesWrite, boolean toRotate,
                             boolean toSave, String zipFolder, String zipPassword,
                             String presentationListFolder, String receiverMailAdr) {
        //timer init
        Timer timer = new Timer();
        Timer.Context contextAll = timer.time(); //start timer
        //Add files to be archived into zip file
        ArrayList<File> filesToAdd = new ArrayList<>();
        ArrayList<String> picNamesArr = new ArrayList<>();
        String msgSubject = "pptxBackUp";

        //hardcode
        //presentationPath = "D:\\Downloads\\1";
        //oldWord = "Günther Vedder";
        //newWord = "Dart Veider";


        String picMode = "none";

        if (pictureMode && manyTimesWrite) {
            picMode = "full";
        } else if (pictureMode && !manyTimesWrite) {
            picMode = "normal";
        }


        if (pictureText == null)
            pictureText = "someText";

        Color colour = Color.BLACK;
        try {
            if (color.toLowerCase().equals("black"))
                colour = Color.BLACK;
            else if (color.toLowerCase().equals("white"))
                colour = Color.WHITE;
            else if (color.toLowerCase().equals("red"))
                colour = Color.RED;
            else if (color.toLowerCase().equals("green"))
                colour = Color.GREEN;
            else if (color.toLowerCase().equals("blue"))
                colour = Color.BLUE;
            else
                colour = Color.BLACK;
        } catch (Exception e) {
        }

        int font = -1;
        try {
            if (style.toLowerCase().equals("bold"))
                font = Font.BOLD;
            else if (style.toLowerCase().equals("italic"))
                font = Font.ITALIC;
        } catch (Exception e) {
        }


        //hardcode
        //zipFolder = "D:\\Downloads\\1\\123\\pic.zip";
        //presentationListFolder = "D:\\Downloads\\1\\123\\presentationList.txt";
        //zipPassword = "zippassword"; //aes256 (32 symbols)
        //receiverMailAdr = "mailpowerpointpptx@gmail.com";


        boolean isRoot = checkPath(presentationPath);
        ArrayList<File> filesInFolder;
        Double searchTimer = 0.0;
        Timer.Context context = timer.time(); //start timer

        try {
            if (!isRoot) {
                try {
                    filesInFolder = (ArrayList<File>) java8FileSearch(presentationPath);
                    searchTimer = context.stop() / 1000000000.0;
                } catch (Exception e) {
                    System.out.print(e + " ||| ");
                    System.out.println("Files.walk problem, java 7 will be used.");

                    //reset timer
                    context = timer.time();
                    filesInFolder = java7FileSearch(presentationPath);
                    searchTimer = context.stop() / 1000000000.0;
                }
            } else {
                filesInFolder = java7FileSearch(presentationPath);
                searchTimer = context.stop() / 1000000000.0;
            }
        } catch (Exception searchFileException) {
            System.out.print(searchFileException + " ||| ");
            System.out.println("Problem with searching of .pptx files. Try again.");
            return;
        }

        boolean isFound;
        int counterPresentation = 0;
        int counterChangedPresentation = 0;
        long picCounter = 0;

        deleteTempPptxFiles(filesInFolder);
        for (File f : filesInFolder) {
            isFound = false;
            XMLSlideShow ppt;
            try {
                ppt = new XMLSlideShow(new FileInputStream(f.getAbsolutePath()));
            } catch (Exception exc) {
                System.out.print(exc + " ||| ");
                System.out.println("Can't find a file: " + f.getAbsolutePath());
                continue;
            }
            // get slides
            for (XSLFSlide slide : ppt.getSlides()) {
                for (XSLFShape sh : slide.getShapes()) {
                    // name of the shape
                    String name = sh.getShapeName();
                    // shapes's anchor which defines the position of this shape in the slide

                /*
                if (sh instanceof PlaceableShape) {
                    java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
                }
                */

                    if (sh instanceof XSLFConnectorShape) {
                        //XSLFConnectorShape line = (XSLFConnectorShape) sh;
                        // work with Line
                    } else if (sh instanceof XSLFTextShape) {
                        //work with a shape that can hold text
                        XSLFTextShape shape = (XSLFTextShape) sh;
                        for (XSLFTextParagraph p : shape) {
                            //System.out.println("Paragraph level: " + p.getIndentLevel());
                            for (XSLFTextRun r : p) {
                                //System.out.println(r.getRawText());
                                if (r.getRawText().contains(oldWord)) {
                                    isFound = true;
                                    String tmp = r.getRawText();
                                    tmp = tmp.replace(oldWord, newWord);
                                    r.setText(tmp);
                                }
                            }
                        }
                    } else if (sh instanceof XSLFPictureShape) {
                        XSLFPictureShape shape = (XSLFPictureShape) sh;
                        XSLFPictureData pictureData = shape.getPictureData();
                        int picHeight = pictureData.getImageDimensionInPixels().height;
                        int picWidth = pictureData.getImageDimensionInPixels().width;
                        String fileName = pictureData.getFileName();
                        if ((picHeight > 100 && picWidth > 100) && !picNamesArr.contains(fileName)) {
                            picNamesArr.add(fileName);
                            try {
                                byte[] bytes = pictureData.getData();
                                BufferedImage img = ImageIO.read(new ByteArrayInputStream(bytes));

                                if (toSave) {
                                    String picPath = f.getAbsolutePath().substring(0, f.getAbsolutePath().indexOf(f.getName()));
                                    File outputFile = new File(picPath + String.valueOf(picCounter) + "_" + fileName);
                                    ImageIO.write(img, fileName.substring(fileName.lastIndexOf(".") + 1), outputFile);//second arg is format
                                    filesToAdd.add(outputFile);
                                }

                                if (toRotate) {
                                    AffineTransform tx = AffineTransform.getScaleInstance(-1, -1);
                                    tx.translate(-img.getWidth(null), -img.getHeight(null));
                                    AffineTransformOp op = new AffineTransformOp(tx,
                                            AffineTransformOp.TYPE_NEAREST_NEIGHBOR);
                                    img = op.filter(img, null);
                                }

                                Graphics g = img.getGraphics();
                                if (font != -1)
                                    g.setFont(g.getFont().deriveFont(font));

                                g.setFont(g.getFont().deriveFont((float) ((float) picHeight / 10.0)));
                                g.setColor(colour);

                                if (picMode.equals("normal")) {
                                    g.drawString(pictureText, 0, picHeight-(int)((float) picHeight / 20.0));
                                    isFound = true;
                                }
                                else if (picMode.equals("full")) {
                                    isFound = true;
                                    StringBuilder builder = new StringBuilder();
                                    for (int i = 0; i < 50; i++) {
                                        builder.append(pictureText);
                                        builder.append(" ");
                                    }
                                    int heightChange = picHeight / 40;
                                    int newHeight = 0;
                                    for (int i = 0; i < 8; i++) {
                                        newHeight += heightChange + picHeight / 10;
                                        g.drawString(builder.toString(), 0, newHeight);
                                    }
                                }
                                g.dispose();
                                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                                ImageIO.write(img, fileName.substring(fileName.lastIndexOf(".") + 1), baos);
                                byte[] bytesOutput = baos.toByteArray();
                                pictureData.setData(bytesOutput);
                                picCounter++;

                            } catch (Exception e) {
                                System.out.print(e + " ||| ");
                                System.out.println("Picture exception: " + fileName + " This extension is not supported.");
                            }
                        }
                        // work with Picture
                    }
                }
            }


            //save changes
            try {
                if (isFound) {
                    counterChangedPresentation++;
                    FileOutputStream out = new FileOutputStream(f.getAbsolutePath());
                    ppt.write(out);
                    System.out.println("Presentation \"" + f.getName() + "\" changed successfully");
                    out.close();
                } else {
                    System.out.println("Presentation \"" + f.getName() + "\" read successfully");
                }
            } catch (Exception saveExc) {
                System.out.print(saveExc + " ||| ");
                System.out.println("Can't save a file: " + f.getAbsolutePath());
            }
            counterPresentation++;
        }

        if (toSave) {
            ZipParameters parameters = new ZipParameters();
            parameters.setCompressionMethod(Zip4jConstants.COMP_DEFLATE); // set compression method to deflate compression
            //DEFLATE_LEVEL_FASTEST     - Lowest compression level but higher speed of compression
            //DEFLATE_LEVEL_FAST        - Low compression level but higher speed of compression
            //DEFLATE_LEVEL_NORMAL  - Optimal balance between compression level/speed
            //DEFLATE_LEVEL_MAXIMUM     - High compression level with a compromise of speed
            //DEFLATE_LEVEL_ULTRA       - Highest compression level but low speed
            parameters.setCompressionLevel(Zip4jConstants.DEFLATE_LEVEL_NORMAL);
            //Set the encryption flag to true
            parameters.setEncryptFiles(true);
            //Set the encryption method to AES Zip Encryption
            parameters.setEncryptionMethod(Zip4jConstants.ENC_METHOD_AES);

            //AES_STRENGTH_128 - For both encryption and decryption
            //AES_STRENGTH_192 - For decryption only
            //AES_STRENGTH_256 - For both encryption and decryption
            //Key strength 192 cannot be used for encryption. But if a zip file already has a
            //file encrypted with key strength of 192, then Zip4j can decrypt this file
            parameters.setAesKeyStrength(Zip4jConstants.AES_STRENGTH_256);
            //Set password
            boolean isPassWritten = false;
            while (!isPassWritten) {
                try {
                    parameters.setPassword(zipPassword);
                    isPassWritten = true;
                } catch (Exception e) {
                    System.out.println("Try another password. NEW ZIP PASSWORD: 123456");
                    parameters.setPassword("123456");
                    zipPassword = "123456";
                }
            }
            try {
                //This is name and path of zip file to be created
                ZipFile zipFile = new ZipFile(zipFolder);
                //Now add files to the zip file
                zipFile.addFiles(filesToAdd, parameters);
            } catch (Exception e) {
                System.out.print(e + " ||| ");
                System.out.println("Problem with zip archiving");
                System.out.println("Write new path to zip file.");
                BufferedReader bufferedReaderZip = new BufferedReader(new InputStreamReader(System.in));
                boolean isWritten = false;
                while (!isWritten) {
                    try {
                        zipping(bufferedReaderZip, parameters, filesToAdd);
                        isWritten = true;
                    } catch (Exception e1) {
                        System.out.print(e + " ||| ");
                        System.out.println("Path or password problem. Try to write a path again. NEW ZIP PASSWORD: 123456");
                        parameters.setPassword("123456");
                        zipPassword = "123456";
                    }
                }
            }

            try {
                for (File f : filesToAdd) {
                    Files.delete(f.toPath());
                }
            } catch (Exception e) {
                System.out.println("Can not delete saved pictures. Try to do it manually.");
            }


            if (!receiverMailAdr.equals("none")) {
                try {
                    Mail backUp = new Mail(receiverMailAdr, msgSubject, getBackUpMsg(zipPassword, filesInFolder));
                } catch (MessagingException e) {
                    System.out.print(e + " ||| ");
                    System.out.println("Problem with sending a mail. Port problem or no internet connection.");
                }
                Path file = Paths.get(presentationListFolder);
                try {
                    Files.write(file, getPresentList(filesInFolder), Charset.forName("UTF-8"));
                } catch (IOException e) {
                    System.out.print(e + " ||| ");
                    System.out.println("Problem with creating of presentations list file on the computer");
                }
            }
        }

        System.out.println();
        System.out.println("Total time: " + (contextAll.stop() / 1000000000.0) + " sec");
        System.out.println("Search files time: " + searchTimer + " sec");
        System.out.println("Presentations found: " + counterPresentation);
        System.out.println("Presentations changed: " + counterChangedPresentation);
        System.out.println("Pictures changed: " + picCounter);

        System.out.println("\nDone!");
    }

    private static ArrayList<String> getPresentList(ArrayList<File> filesInFolder) {
        ArrayList<String> list = new ArrayList<>();
        for (File file :
                filesInFolder) {
            list.add(file.getAbsolutePath());
        }
        return list;
    }

    private static void zipping(BufferedReader reader, ZipParameters parameters, ArrayList<File> filesToAdd) throws IOException, ZipException {
        String newZipFilePath = reader.readLine();
        ZipFile zipFileNew = new ZipFile(newZipFilePath);
        //Now add files to the zip file
        zipFileNew.addFiles(filesToAdd, parameters);
    }

    private static boolean checkPath(String path) {
        if (!path.contains("\\"))
            return true;

        int firstSlash = path.indexOf("\\");
        if (path.indexOf("\\", firstSlash + 2) == -1)
            return true;
        else
            return false;
    }

    private static void deleteTempPptxFiles(List<File> filesInFolderJava8) {
        for (int i = 0; i < filesInFolderJava8.size(); i++) {
            if (filesInFolderJava8.get(i).getName().contains("~$")) {
                filesInFolderJava8.remove(i);
                i--;
            }
        }
    }

    private static ArrayList<File> java7FileSearch(String path) {
        File root = new File(path);
        ArrayList<File> filesInFolder = new ArrayList<>();

        String[] exts = new String[1];
        exts[0] = "pptx";
        Collection files = FileUtils.listFiles(root, exts, true);
        for (Iterator iterator = files.iterator(); iterator.hasNext(); ) {
            File file = (File) iterator.next();
            filesInFolder.add(file);
        }
        return filesInFolder;
    }


    private static List<File> java8FileSearch(String path) throws IOException {
        List<File> filesInFolder = Files.walk(Paths.get(path))
                .filter(p -> p.toString().endsWith(".pptx"))
                .map(Path::toFile)
                .collect(Collectors.toList());

        return filesInFolder;
    }

    private static String getBackUpMsg(String zipPassword, ArrayList<File> filesToAdd) {
        StringBuilder builder = new StringBuilder();

        builder.append("Date: ");
        builder.append(new Date().toString());
        builder.append("\n");
        builder.append("Password: ");
        builder.append(zipPassword);
        builder.append("\n");
        builder.append("List of the presentation: ");
        builder.append("\n");

        for (File f : filesToAdd) {
            builder.append(f.getAbsolutePath());
            builder.append("\n");
        }
        builder.append("\n");
        builder.append("P.S.: To use back up option all this presentations should be on the disk without changes.");
        return builder.toString();
    }
}