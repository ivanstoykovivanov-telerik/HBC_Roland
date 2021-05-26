// ****************************************************************************************************
// *
// *   ARIS 7.1 - SpiderView.rsm
// *   Calculate a graph of assigned follow-up processes.
// *
// *   (c) Bayrische Hypo- und Vereinsbank AG, ein Unternehmen der Unicredit Group
// *
// ****************************************************************************************************
// Version 1.0.0 (2014-06-03|Steffen Ploetz)
// Initial revision
// ****************************************************************************************************
// Version 1.1.0 (2014-08-11|Steffen Ploetz)
// Extended to process interfaces.
// ****************************************************************************************************



/*
// ############################################################################################
// ############################################################################################
// #### First EMF appempt - without any success!
// #### ARIS doesn't evaluate WMF / EMF image size.
// #### Only the left top excerpt of about 32 mm x 11 mm will be displayed.
// ############################################################################################
// ############################################################################################

// ############################################################################################
// Create EMF from several parts, concat them into a byte buffer and create graphic from buffer.
var emfLeadIn  = [0x01, 0x00, 0x00, 0x00,  0x6C, 0x00, 0x00, 0x00,    0xDC, 0xFF, 0xFF, 0xFF,  0x05, 0x00, 0x00, 0x00,
                  0x88, 0x00, 0x00, 0x00,  0x91, 0x00, 0x00, 0x00,    0x0B, 0xFB, 0xFF, 0xFF,  0xB0, 0x00, 0x00, 0x00,
                  0xBB, 0x12, 0x00, 0x00,  0xFB, 0x13, 0x00, 0x00,    0x20, 0x45, 0x4D, 0x46,  0x00, 0x00, 0x01, 0x00,
                  0xB4, 0x02, 0x00, 0x00,  0x14, 0x00, 0x00, 0x00,    0x02, 0x00, 0x00, 0x00,  0x00, 0x00, 0x00, 0x00,
                  0x00, 0x00, 0x00, 0x00,  0x00, 0x00, 0x00, 0x00,    0x80, 0x07, 0x00, 0x00,  0x38, 0x04, 0x00, 0x00,
                  0xA5, 0x02, 0x00, 0x00,  0x7D, 0x01, 0x00, 0x00,    0x00, 0x00, 0x00, 0x00,  0x00, 0x00, 0x00, 0x00,
                  0x00, 0x00, 0x00, 0x00,  0xD5, 0x55, 0x0A, 0x00,    0x48, 0xD0, 0x05, 0x00,  0x46, 0x00, 0x00, 0x00,
                  0x2C, 0x00, 0x00, 0x00,  0x20, 0x00, 0x00, 0x00,    0x45, 0x4D, 0x46, 0x2B,  0x01, 0x40, 0x01, 0x00,
                  0x1C, 0x00, 0x00, 0x00,  0x10, 0x00, 0x00, 0x00,    0x02, 0x10, 0xC0, 0xDB,  0x01, 0x00, 0x00, 0x00,
                  0x60, 0x00, 0x00, 0x00,  0x60, 0x00, 0x00, 0x00,    0x46, 0x00, 0x00, 0x00,  0x5C, 0x00, 0x00, 0x00,
                  0x50, 0x00, 0x00, 0x00,  0x45, 0x4D, 0x46, 0x2B];
var emfLeadOut = [0x02, 0x40, 0x00, 0x00,  0x0C, 0x00, 0x00, 0x00,    0x00, 0x00, 0x00, 0x00,  0x0E, 0x00, 0x00, 0x00,
                  0x14, 0x00, 0x00, 0x00,  0x00, 0x00, 0x00, 0x00,    0x10, 0x00, 0x00, 0x00,  0x14, 0x00, 0x00, 0x00];

var fullEmfBytes    = java.nio.ByteBuffer.allocate(emfLeadIn.length + emfLeadOut.length);
for (var countBytes = 0; countBytes < emfLeadIn.length; countBytes++)
{
    if (emfLeadIn[countBytes] < 128)
        fullEmfBytes.put(emfLeadIn[countBytes]);
    else
        fullEmfBytes.put(emfLeadIn[countBytes] - 256);
}
for (var countBytes = 0; countBytes < emfLeadOut.length; countBytes++)
{
    if (emfLeadOut[countBytes] < 128)
        fullEmfBytes.put(emfLeadOut[countBytes]);
    else
        fullEmfBytes.put(emfLeadOut[countBytes] - 256);
}
var logoA = Context.createPicture(fullEmfBytes.array(), Constants.IMAGE_FORMAT_EMF);
p_output.OutGraphic(logoA, -1, 6, 4);
*/
/*
// ############################################################################################
// ############################################################################################
// #### Second EMF appempt - without any success!
// #### ARIS doesn't evaluate WMF / EMF image size.
// #### Only the left top excerpt of about 32 mm x 11 mm will be displayed.
// ############################################################################################
// ############################################################################################

// ############################################################################################
// Load EMF file into buffer and create graphic from buffer.
var emfFile = new java.io.File("C:\\TEMP\\LogoL.emf");
var emfFileStream = new java.io.FileInputStream(emfFile);
var emfFileBuffer = java.lang.reflect.Array.newInstance(java.lang.Byte.TYPE, emfFile.length());
with (new java.io.BufferedInputStream(emfFileStream) )
{
    read(emfFileBuffer);
    close();
}
var logoB = Context.createPicture(emfFileBuffer, Constants.IMAGE_FORMAT_EMF);
p_output.OutGraphic(logoB, -1, 6, 4);
*/

/*
// ############################################################################################
// ############################################################################################
// #### Counter-check PNG - with success!
// #### ARIS correctly evaluates bitmap image size.
// #### The full image will be displayed.
// ############################################################################################
// ############################################################################################

// ############################################################################################
// Create PNG from a java script byte array, copy them into a byte buffer and create graphic from buffer.

var gifEvtSmall = [ ... ];
var gifFktSmall = [ ... ];

var fullEvtBytes    = java.nio.ByteBuffer.allocate(gifEvtSmall.length);
for (var countBytes = 0; countBytes < gifEvtSmall.length; countBytes++)
{
    if (gifEvtSmall[countBytes] < 128)
        fullEvtBytes.put(gifEvtSmall[countBytes]);
    else
        fullEvtBytes.put(gifEvtSmall[countBytes] - 256);
}
var logo4 = Context.createPicture(fullEvtBytes.array(), Constants.IMAGE_FORMAT_PNG);

var fullFktBytes    = java.nio.ByteBuffer.allocate(gifFktSmall.length);
for (var countBytes = 0; countBytes < gifFktSmall.length; countBytes++)
{
    if (gifFktSmall[countBytes] < 128)
        fullFktBytes.put(gifFktSmall[countBytes]);
    else
        fullFktBytes.put(gifFktSmall[countBytes] - 256);
}
var logoC = Context.createPicture(fullFktBytes.array(), Constants.IMAGE_FORMAT_PNG);
p_output.OutGraphic(logoC, -1, 6, 4);
*/


// ############################################################################################
// ############################################################################################
// #### Final code attempt.
// #### Use native HSSF to create XLS output and serialize it into an ARIS-ExcelWorkbook. 
// #### Everything works fine.
// ############################################################################################
// ############################################################################################

// ######################################################################
// Customer properties.
// ######################################################################

var iMaxEvaluationDepth = 254;

// ######################################################################
// Constant values.
// ######################################################################

var iHAlignLeft   = 1;
var iHAlignCenter = 2;
var iVAlignTop    = 0;
var iVAlignMiddle = 1;
var iVAlignBottom = 2;

var pngEvtSmall = [ 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, 
                    0x00, 0x00, 0x00, 0x13, 0x00, 0x00, 0x00, 0x0C, 0x08, 0x06, 0x00, 0x00, 0x00, 0x80, 0xD0, 0x86, 
                    0x82, 0x00, 0x00, 0x00, 0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xAE, 0xCE, 0x1C, 0xE9, 0x00, 0x00, 
                    0x00, 0x04, 0x67, 0x41, 0x4D, 0x41, 0x00, 0x00, 0xB1, 0x8F, 0x0B, 0xFC, 0x61, 0x05, 0x00, 0x00, 
                    0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00, 0x0E, 0xC4, 0x00, 0x00, 0x0E, 0xC4, 0x01, 0x95, 
                    0x2B, 0x0E, 0x1B, 0x00, 0x00, 0x01, 0xA2, 0x49, 0x44, 0x41, 0x54, 0x38, 0x4F, 0x8D, 0x91, 0x59, 
                    0x4F, 0xC2, 0x50, 0x10, 0x46, 0xFB, 0xFF, 0xDF, 0x5C, 0x02, 0xA2, 0x22, 0xCA, 0xD2, 0x96, 0xD2, 
                    0x52, 0x36, 0xAB, 0xA2, 0x46, 0x13, 0x4D, 0x10, 0xD0, 0x40, 0x5C, 0x22, 0x31, 0xC6, 0x85, 0x52, 
                    0x84, 0xB6, 0xE2, 0x1A, 0xA3, 0xF1, 0xE1, 0x73, 0x6E, 0x6F, 0x11, 0x30, 0x28, 0xDE, 0xE4, 0xB4, 
                    0x69, 0x32, 0x3D, 0xF7, 0x9B, 0x19, 0x01, 0xFE, 0x29, 0x47, 0x62, 0xA8, 0xCE, 0xCE, 0xFB, 0x84, 
                    0x50, 0x9D, 0xE9, 0x33, 0x87, 0x0A, 0x63, 0xBA, 0x4F, 0x90, 0x33, 0x15, 0x44, 0x79, 0x2A, 0x80, 
                    0x83, 0xF9, 0x88, 0x6F, 0x00, 0x3C, 0x59, 0xCF, 0x6C, 0xE1, 0x54, 0xD6, 0xE0, 0xAE, 0x16, 0xE1, 
                    0x16, 0x8A, 0x70, 0x0A, 0x1B, 0x70, 0xF2, 0x44, 0x6E, 0x1D, 0x76, 0x6E, 0x0D, 0x76, 0x76, 0x0D, 
                    0xDD, 0x8C, 0xC1, 0xD1, 0x57, 0xD1, 0xD1, 0x0B, 0xE8, 0xA4, 0x0B, 0xB8, 0xD3, 0xF2, 0x38, 0x59, 
                    0x49, 0xE0, 0xFE, 0xA6, 0x39, 0x90, 0xD5, 0xD4, 0x0C, 0x3A, 0x24, 0x9A, 0x2C, 0x21, 0x48, 0xD2, 
                    0x49, 0xE7, 0x3D, 0xD1, 0x9D, 0x9A, 0x83, 0x99, 0xD4, 0x71, 0x18, 0x97, 0xB9, 0xEC, 0xF3, 0xFD, 
                    0x1D, 0x95, 0x70, 0x74, 0x44, 0xE2, 0xE8, 0x06, 0xDC, 0x54, 0x16, 0xAE, 0xC2, 0x71, 0x92, 0x84, 
                    0x9C, 0x21, 0x74, 0xD8, 0x44, 0x97, 0x68, 0x93, 0xA8, 0x4D, 0x35, 0x6D, 0x25, 0x83, 0x83, 0xC0, 
                    0x22, 0x3E, 0xDE, 0xDE, 0x20, 0x5C, 0x95, 0xCA, 0xB8, 0xA6, 0x9B, 0x6C, 0x26, 0xA2, 0x24, 0x76, 
                    0xD6, 0xC0, 0xE3, 0x72, 0x12, 0x2F, 0xA1, 0xD8, 0x9F, 0x74, 0x45, 0xCD, 0x13, 0x59, 0x94, 0xEC, 
                    0x62, 0x59, 0xC4, 0xC5, 0xEE, 0x1E, 0x84, 0x12, 0xA5, 0xEA, 0x4B, 0x58, 0x3B, 0x2E, 0x15, 0x8C, 
                    0xFB, 0xF9, 0x27, 0x4F, 0x0B, 0x71, 0x58, 0x94, 0xB0, 0x25, 0xA7, 0x3D, 0xF6, 0x29, 0x9D, 0xD0, 
                    0xD8, 0xDA, 0x41, 0x93, 0xE2, 0x7E, 0xCF, 0x85, 0x52, 0x3E, 0x53, 0xE1, 0x38, 0xC1, 0x30, 0xBD, 
                    0x88, 0xC4, 0x45, 0x92, 0x86, 0xCB, 0x88, 0x88, 0x33, 0x63, 0x03, 0xC2, 0xAB, 0x7B, 0x8F, 0x7A, 
                    0x54, 0x1A, 0xD9, 0x90, 0x43, 0x2D, 0x3C, 0x2E, 0x49, 0x3E, 0xA2, 0xC7, 0x03, 0x23, 0xCC, 0xE9, 
                    0x11, 0x96, 0xA8, 0xC2, 0xA4, 0x3A, 0x93, 0xDE, 0x47, 0xA1, 0x25, 0xBC, 0xD8, 0x0E, 0xDF, 0x66, 
                    0x25, 0x2A, 0x7A, 0xB2, 0xFE, 0x86, 0x86, 0x87, 0x6B, 0xF9, 0x73, 0xB1, 0xFC, 0x76, 0x5A, 0x12, 
                    0x4F, 0xC3, 0x24, 0x66, 0x42, 0x45, 0x33, 0x91, 0x42, 0x29, 0x14, 0xE6, 0xDB, 0x64, 0x8F, 0xC6, 
                    0xF6, 0x2E, 0x5A, 0xC6, 0xE6, 0x78, 0x09, 0xF1, 0x9B, 0xA4, 0x19, 0x57, 0x70, 0x45, 0xDF, 0xE7, 
                    0xC5, 0xED, 0x81, 0x8C, 0xB5, 0x7A, 0xAC, 0xA4, 0x07, 0xA2, 0x89, 0x12, 0x2E, 0xBA, 0x8D, 0x29, 
                    0xA8, 0x51, 0x57, 0xAC, 0xC5, 0x6F, 0x19, 0x3B, 0x87, 0xB2, 0x8A, 0x7A, 0x52, 0xFB, 0x17, 0x35, 
                    0x0F, 0xD5, 0xA3, 0x12, 0x93, 0x7C, 0x03, 0xF0, 0x05, 0x64, 0x8B, 0xED, 0x2F, 0x17, 0xA8, 0xA3, 
                    0xFE, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82];
var pngFktSmall = [ 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, 
                    0x00, 0x00, 0x00, 0x13, 0x00, 0x00, 0x00, 0x0C, 0x08, 0x06, 0x00, 0x00, 0x00, 0x80, 0xD0, 0x86, 
                    0x82, 0x00, 0x00, 0x00, 0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xAE, 0xCE, 0x1C, 0xE9, 0x00, 0x00, 
                    0x00, 0x04, 0x67, 0x41, 0x4D, 0x41, 0x00, 0x00, 0xB1, 0x8F, 0x0B, 0xFC, 0x61, 0x05, 0x00, 0x00, 
                    0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00, 0x0E, 0xC4, 0x00, 0x00, 0x0E, 0xC4, 0x01, 0x95, 
                    0x2B, 0x0E, 0x1B, 0x00, 0x00, 0x01, 0x73, 0x49, 0x44, 0x41, 0x54, 0x38, 0x4F, 0x8D, 0x53, 0xC9, 
                    0x4E, 0x02, 0x41, 0x10, 0x7D, 0xB3, 0x20, 0x0C, 0x0C, 0x3A, 0x22, 0x5C, 0xE5, 0x2B, 0x8C, 0x57, 
                    0x13, 0x13, 0x13, 0xBF, 0xD7, 0x84, 0x68, 0x4C, 0x4C, 0x4C, 0x4C, 0x38, 0xF9, 0x1D, 0x72, 0x32, 
                    0x4A, 0x18, 0x66, 0x63, 0xBA, 0xAC, 0xAA, 0xEE, 0x66, 0x51, 0x0F, 0x14, 0xBC, 0xAE, 0xAA, 0xD7, 
                    0xB5, 0xCE, 0x40, 0xD0, 0xB6, 0x2D, 0xBD, 0xBE, 0xBF, 0x60, 0xB5, 0x5E, 0x22, 0x0C, 0x63, 0x58, 
                    0x21, 0xA7, 0xF7, 0x2D, 0x11, 0xF6, 0xF6, 0x08, 0x31, 0x89, 0x0C, 0xFA, 0xDD, 0x01, 0x6E, 0xAE, 
                    0x6E, 0x11, 0x3C, 0xCF, 0x1F, 0x69, 0x32, 0x1E, 0x23, 0xE9, 0xF7, 0xF9, 0x82, 0x2C, 0x5C, 0xC6, 
                    0x81, 0xAF, 0xDA, 0x71, 0xDE, 0x77, 0x76, 0xD3, 0x34, 0xF8, 0xFE, 0xE4, 0x61, 0xAA, 0xBA, 0x44, 
                    0x2F, 0x49, 0x60, 0xB8, 0x83, 0xBF, 0x14, 0x6D, 0x8C, 0x71, 0x9C, 0xE5, 0x8D, 0xE2, 0xD0, 0xF7, 
                    0xB1, 0x51, 0x14, 0xA1, 0xA8, 0x73, 0xDE, 0x8C, 0x0D, 0x2D, 0xE2, 0x82, 0xFF, 0x2B, 0x62, 0xED, 
                    0xBF, 0xBE, 0x8F, 0x15, 0x00, 0x21, 0x82, 0xD9, 0xDB, 0x03, 0x4D, 0x2F, 0xA7, 0xB2, 0x00, 0x07, 
                    0x00, 0x85, 0xC9, 0xB1, 0x32, 0x5F, 0xEC, 0x89, 0xD8, 0x53, 0xC5, 0x99, 0x1C, 0xA5, 0x3A, 0xE0, 
                    0xE4, 0x94, 0x32, 0xE5, 0xA5, 0xC9, 0xE2, 0x63, 0xC1, 0x8C, 0x5C, 0xBB, 0x8E, 0xD2, 0xA1, 0x32, 
                    0x25, 0x2A, 0x2A, 0x51, 0x2B, 0xAA, 0x1D, 0x60, 0xD1, 0xF0, 0x29, 0xA8, 0x51, 0xA2, 0xA5, 0x8D, 
                    0xE6, 0xC8, 0x94, 0x22, 0xA1, 0x54, 0xB6, 0x84, 0x2D, 0xBA, 0x1D, 0xE1, 0x08, 0x91, 0x29, 0xFD, 
                    0x20, 0x22, 0x3C, 0x99, 0x1D, 0xD3, 0x77, 0x38, 0xBE, 0x14, 0x0F, 0x61, 0xA4, 0x98, 0xCF, 0x23, 
                    0x84, 0x3A, 0x8B, 0x9B, 0x4C, 0x3A, 0x0C, 0x90, 0x22, 0xC3, 0x04, 0x67, 0x34, 0xC6, 0x29, 0x5D, 
                    0x60, 0x68, 0x46, 0x5B, 0xA4, 0xE6, 0x1C, 0x83, 0x36, 0x53, 0xA4, 0x9B, 0x0C, 0x01, 0x05, 0x9A, 
                    0xA3, 0x1B, 0xF1, 0x37, 0x6C, 0x5B, 0xD9, 0x7B, 0xD7, 0x41, 0xBA, 0x45, 0x14, 0x2B, 0x62, 0xEA, 
                    0x28, 0xD4, 0x37, 0x11, 0x43, 0x74, 0x8C, 0x90, 0x6D, 0x18, 0x7E, 0x42, 0xFC, 0x36, 0x35, 0xC7, 
                    0xE5, 0x86, 0xBD, 0x93, 0x04, 0xF9, 0x7A, 0xA5, 0xAF, 0x59, 0x3A, 0xE8, 0xC7, 0x75, 0xDB, 0xAE, 
                    0xAE, 0xB6, 0xF5, 0x95, 0xFB, 0x15, 0x5B, 0x14, 0x6B, 0x74, 0x3B, 0x3D, 0xE8, 0xDF, 0xE9, 0x69, 
                    0x3E, 0xC3, 0x32, 0xF7, 0x3F, 0x07, 0x11, 0x6B, 0xE9, 0xC9, 0xC1, 0x5E, 0xF6, 0x58, 0x47, 0x4B, 
                    0x31, 0x83, 0x6C, 0x38, 0xC2, 0xDD, 0xF5, 0x3D, 0x7E, 0x00, 0x6E, 0x37, 0x95, 0xC9, 0x0A, 0xE9, 
                    0x3E, 0xA8, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82];
var pngIFaSmall = [ 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,   0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
                    0x00, 0x00, 0x00, 0x13, 0x00, 0x00, 0x00, 0x0C,   0x08, 0x06, 0x00, 0x00, 0x00, 0x80, 0xD0, 0x86,
                    0x82, 0x00, 0x00, 0x00, 0x01, 0x73, 0x52, 0x47,   0x42, 0x00, 0xAE, 0xCE, 0x1C, 0xE9, 0x00, 0x00,
                    0x00, 0x04, 0x67, 0x41, 0x4D, 0x41, 0x00, 0x00,   0xB1, 0x8F, 0x0B, 0xFC, 0x61, 0x05, 0x00, 0x00,
                    0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00,   0x0E, 0xC4, 0x00, 0x00, 0x0E, 0xC4, 0x01, 0x95,
                    0x2B, 0x0E, 0x1B, 0x00, 0x00, 0x00, 0xA8, 0x49,   0x44, 0x41, 0x54, 0x38, 0x4F, 0xAD, 0x92, 0xD1,
                    0x0D, 0x83, 0x20, 0x14, 0x45, 0x9F, 0xD4, 0xB6,   0x5B, 0xF8, 0xC9, 0x14, 0x6C, 0xD2, 0x56, 0x13,
                    0xF6, 0x82, 0xB0, 0x8B, 0x4C, 0xC1, 0x27, 0x7B,   0x10, 0x9A, 0x4B, 0x5F, 0x8C, 0x1A, 0x93, 0x62,
                    0xE9, 0xF9, 0x50, 0x34, 0x8F, 0x13, 0xAE, 0xDE,   0x2E, 0xA5, 0x94, 0x9D, 0x73, 0x14, 0x63, 0xA4,
                    0x5A, 0x86, 0x61, 0xA0, 0x71, 0x1A, 0xA9, 0xBF,   0xF4, 0xFC, 0x86, 0x31, 0xC6, 0xE4, 0x79, 0x9E,
                    0xF3, 0x19, 0x30, 0x8F, 0x7D, 0x7B, 0x04, 0x4E,   0xA4, 0x94, 0x62, 0x75, 0x1D, 0x98, 0x3F, 0x4A,
                    0x22, 0xF8, 0xFE, 0x17, 0x9A, 0x64, 0xD6, 0x5A,   0x5E, 0x7D, 0x68, 0x92, 0x49, 0x29, 0x37, 0xC2,
                    0x26, 0x19, 0xBE, 0xDD, 0x5A, 0x28, 0xF0, 0x9B,   0xBD, 0xF7, 0xE5, 0xA1, 0x16, 0xCC, 0x63, 0x1F,
                    0x58, 0x0B, 0x7F, 0xEE, 0xD9, 0xF3, 0xF5, 0xA0,   0xDB, 0xF5, 0x5E, 0xC4, 0x21, 0x04, 0xD2, 0x5A,
                    0x13, 0x71, 0x45, 0x0A, 0x67, 0x3B, 0xB7, 0xEF,   0xDB, 0x46, 0x06, 0x6A, 0x85, 0x47, 0xC5, 0xED,
                    0x70, 0xE1, 0xD3, 0x2F, 0x20, 0xFF, 0xB7, 0xD8,   0x88, 0x5A, 0xA2, 0x2D, 0x10, 0xBD, 0x01, 0x3A,
                    0x1B, 0x12, 0x39, 0xF8, 0xFE, 0xFD, 0x9A, 0x00,   0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
                    0x42, 0x60, 0x82];

// ######################################################################
// User defined data types.
// ######################################################################

// HierarchyRecord
// <summary>
// This structure holds the hierarchy information
// associated with a level 5 model (normally of type MT_EEPC).
// </summary>
__usertype_processrecord = function()
{
    this.predecessor = null;            // The predecessor as __usertype_processrecord,
    this.startEvent = null;             // The ARIS obj occ of the event, that invokes this process.
    this.process = null;                // The ARIS model, that represets the node.
    this.deadend = null;                // The ARIS object, that represets an interface without assignments.
    this.successors = new Array();      // The cuccessors as array of __usertype_processrecord,
    this.followupDepth = 0;             // The number of follow-up levels.
    this.requiredRows  = 0;             // The number of rows required to draw this graph node.
    this.startRow = -1;                 // The zero-based start row index.
}

// ######################################################################
// Global variables.
// ######################################################################

var nLocale = 1031;
var oHssfCellStyle = null;
var oDoneProcessList = new Array();

// ######################################################################
// Flow control.
// ######################################################################

/// <summary> Caller for main flow control. </summary>
/// <returns> NOT EVALUATED BY CALLER. </returns>
main();

/// <summary> Main flow control. </summary>
/// <returns> NOT EVALUATED BY CALLER. </returns>
function main()
{
    nLocale = Context.getSelectedLanguage();

    // Prepare workbook.
    var oHssfWorkbook = new org.apache.poi.hssf.usermodel.HSSFWorkbook();
    oHssfCellStyle = oHssfWorkbook.createCellStyle();
    oHssfCellStyle.setWrapText(true);
    oHssfCellStyle.setAlignment(iHAlignCenter);
    oHssfCellStyle.setVerticalAlignment(iVAlignTop);

    // Prepare event image to embedd.
    var fullEvtBytes    = java.nio.ByteBuffer.allocate(pngEvtSmall.length);
    for (var countBytes = 0; countBytes < pngEvtSmall.length; countBytes++)
    {
        if (pngEvtSmall[countBytes] < 128)
            fullEvtBytes.put(pngEvtSmall[countBytes]);
        else
            fullEvtBytes.put(pngEvtSmall[countBytes] - 256);
    }
    var evtPictureIdx = oHssfWorkbook.addPicture(fullEvtBytes.array(), org.apache.poi.hssf.usermodel.HSSFWorkbook.PICTURE_TYPE_PNG);
    
    // Prepare function image to embedd.
    var fullFktBytes    = java.nio.ByteBuffer.allocate(pngFktSmall.length);
    for (var countBytes = 0; countBytes < pngFktSmall.length; countBytes++)
    {
        if (pngFktSmall[countBytes] < 128)
            fullFktBytes.put(pngFktSmall[countBytes]);
        else
            fullFktBytes.put(pngFktSmall[countBytes] - 256);
    }
    var fktPictureIdx = oHssfWorkbook.addPicture(fullFktBytes.array(), org.apache.poi.hssf.usermodel.HSSFWorkbook.PICTURE_TYPE_PNG);
    
    // Prepare interface image to embedd.
    var fullIFaBytes    = java.nio.ByteBuffer.allocate(pngIFaSmall.length);
    for (var countBytes = 0; countBytes < pngIFaSmall.length; countBytes++)
    {
        if (pngIFaSmall[countBytes] < 128)
            fullIFaBytes.put(pngIFaSmall[countBytes]);
        else
            fullIFaBytes.put(pngIFaSmall[countBytes] - 256);
    }
    var ifaPictureIdx = oHssfWorkbook.addPicture(fullIFaBytes.array(), org.apache.poi.hssf.usermodel.HSSFWorkbook.PICTURE_TYPE_PNG);

    // Collect models to evaluate.
    var oSelectedModels = ArisData.getSelectedModels();
    var oModelsToProcess = new Array();
    
    if (oSelectedModels.length == 0)
        return;

    for (var countModels = 0; countModels < oSelectedModels.length; countModels++)
    {
        if (oSelectedModels[countModels].TypeNum() == oSelectedModels[countModels].Database().ActiveFilter().UserDefinedModelTypeNum("970c9df0-9353-11e0-44a4-00300571cf1f"))
        {
            var oDetailProcessSummaryFunctions = oSelectedModels[countModels].ObjDefListFilter(Constants.OT_FUNC);
            for (var countDetailProcessSummaryFunctions = 0; countDetailProcessSummaryFunctions < oDetailProcessSummaryFunctions.length; countDetailProcessSummaryFunctions++)
            {
                var oAssignments = oDetailProcessSummaryFunctions[countDetailProcessSummaryFunctions].AssignedModels(Constants.MT_EEPC);
                for (var countAssignments = 0; countAssignments < oAssignments.length; countAssignments++)
                    oModelsToProcess.push(oAssignments[countAssignments]);
            }
        }
        else if (oSelectedModels[countModels].TypeNum() == Constants.MT_EEPC)
            oModelsToProcess.push(oSelectedModels[countModels]);
    }
    
    if (oModelsToProcess.length == 0)
        return;

    var parseOk = true;
    do
    {    
        var sValue = Dialogs.InputBox("Enter the maximum evaluation depth (must be smaller than 255).", "Limit the output size", iMaxEvaluationDepth.toString());
        if (!isNaN(sValue))
        {
            iMaxEvaluationDepth = parseInt(sValue);
            parseOk = true;
        }
        else
            parseOk = false;
    }
    while (!parseOk || iMaxEvaluationDepth > 254);

    var oHssfPatriarch = null;

    var tProcessRecord = null;
    for (var countModels = 0; countModels < oModelsToProcess.length; countModels++)
    {
        // Evaluate process hierarchy.
        oDoneProcessList = new Array();
        tProcessRecord = evaluateProcess(null, oModelsToProcess[countModels], null, 1);
    
        // Prepare Worksheet.
        var oHssfWorksheet = oHssfWorkbook.createSheet("SpiderView_" + countModels.toString());
        
        // See: http://apache-poi.1045710.n5.nabble.com/Question-about-HSSFSheet-drawing-patriarch-td4555959.html >> post 4ff
        oHssfPatriarch = oHssfWorksheet.createDrawingPatriarch();
        
        // Create all rows, that are required to print out the complete spider view.
        var iMaxRows = tProcessRecord.requiredRows;
        for (var countRows = 0; countRows < iMaxRows; countRows++)
        {
            var row = oHssfWorksheet.createRow(countRows);
            row.setHeightInPoints(8 * oHssfWorksheet.getDefaultRowHeightInPoints());
            
            /* // Not required any longer. Cells will be created directly befor usage.
            for (countCells = 0; countCells < 256; countCells++)
                createCell(oHssfWorksheet, row, countCells);
            */
        }
        // Iterate through evaluation results.
        var iProcessRow = parseInt(((tProcessRecord.requiredRows - 1) / 2).toString());
        printoutProcess(oHssfWorksheet, oHssfPatriarch, tProcessRecord, 0, 0, fktPictureIdx, evtPictureIdx, ifaPictureIdx);
    }

    /* // Not required any longer, since the ByteArrayOutputStream works!
    var fileNode = java.io.File.createTempFile("tmp", ".xls");
    var fileOut = new java.io.FileOutputStream(fileNode.getPath());
    oHssfWorkbook.write(fileOut);
    fileOut.close();
    */
    
    var memoryOut = new java.io.ByteArrayOutputStream();
    oHssfWorkbook.write(memoryOut);
    
    var oWorkbook = Context.createExcelWorkbook(Context.getSelectedFile(), memoryOut.toByteArray());
    var oWorksheet = oWorkbook.createSheet(getString("ID_SHEETNAME"));

    oWorkbook.write();
    oWorksheet = null;
    oWorkbook = null;
}

/// <summary> Iterate through process hierarchy recusively and create a __usertype_processrecord for every hierarchy node. </summary>
/// <param value="tPredecessorRecord"> The {__usertype_processrecord} parent node of the hierarchy node to process/create now. </param>
/// <param value="oStartProcess"> The {ARIS.Model} model representing the hierarchy node to process/create now. </param>
/// <param value="oStartEvent"> The {ARIS.ObjOcc} event that triggered the hierarchy node to process/create now. </param>
/// <param value="iCurrentDepth"> The {int} current hierarchy depth. </param>
/// <returns> The{__usertype_processrecord} created hierarchy node. </returns>
function evaluateProcess(tPredecessorRecord, oStartProcess, oStartEvent, iCurrentDepth)
{
    var tProcessRecord = new __usertype_processrecord();
    tProcessRecord.predecessor = tPredecessorRecord;
    tProcessRecord.startEvent = oStartEvent;
    tProcessRecord.process = oStartProcess;
    tProcessRecord.deadend = null;
    tProcessRecord.requiredRows = 1; // Needed on it's own.
    
    Context.writeStatus("Depth " + iCurrentDepth.toString() + ". Evaluate process '" + oStartProcess.Name(nLocale) + "'.");
    
    // 1. Check, is children have already been processed.
    if (oDoneProcessList.indexOf(oStartProcess.GUID()) >= 0)
        return tProcessRecord;
    else
        oDoneProcessList.push(oStartProcess.GUID());
    
    // 2. Collect all events, that are end events or that trigger a new process.
    var oEventOccList = oStartProcess.ObjOccListFilter(Constants.OT_EVT);
    var oEventOccListToProcess = new Array();
    
    for (var countEvents = 0; countEvents < oEventOccList.length; countEvents++)
    {
        var oCurrentEventOcc = oEventOccList[countEvents];
        // Is current event occurency an end event?
        if (oCurrentEventOcc.Cxns(Constants.EDGES_OUT).length == 0)
        {   ;   }
        // Is current event occurency an event that triggers an interface?
        else if (oCurrentEventOcc.Cxns(Constants.EDGES_OUT).length == 1 && oCurrentEventOcc.Cxns(Constants.EDGES_OUT)[0].TargetObjOcc().OrgSymbolNum() == Constants.ST_PRCS_IF)
        {   ;   }
        // Is current event occurency none of this.
        else
            oCurrentEventOcc = null;
        
        if (oCurrentEventOcc == null)
            continue;
        else
            oEventOccListToProcess.push(oCurrentEventOcc);
    }
    
    // 3. Evaluate successor process.
    var oEvaluatedProcessList = new Array();
    for (var countEvents = 0; countEvents < oEventOccListToProcess.length; countEvents++)
    {
        var oCurrentEventOcc = oEventOccListToProcess[countEvents];
        // Process interface without assigned process.
        if (oCurrentEventOcc.Cxns(Constants.EDGES_OUT).length == 1 &&
            oCurrentEventOcc.Cxns(Constants.EDGES_OUT)[0].TargetObjOcc().OrgSymbolNum() == Constants.ST_PRCS_IF &&
            oCurrentEventOcc.Cxns(Constants.EDGES_OUT)[0].TargetObjOcc().ObjDef().AssignedModels().length == 0)
        {
            tProcessRecord.followupDepth = 1;
            
            var tSubProcessRecord = new __usertype_processrecord();
            tSubProcessRecord.predecessor = tProcessRecord;
            tSubProcessRecord.startEvent = oCurrentEventOcc;
            tSubProcessRecord.process = null;
            tSubProcessRecord.deadend = oCurrentEventOcc.Cxns(Constants.EDGES_OUT)[0].TargetObjOcc().ObjDef();
            tSubProcessRecord.requiredRows = 1; // Needed on it's own.
            
            tProcessRecord.successors.push(tSubProcessRecord);
            tProcessRecord.followupDepth = (tProcessRecord.followupDepth > tSubProcessRecord.followupDepth + 1 ? tProcessRecord.followupDepth : tSubProcessRecord.followupDepth + 1);
            
            if (tProcessRecord.successors.length == 1)
                tProcessRecord.requiredRows  = tSubProcessRecord.requiredRows;
            else
                tProcessRecord.requiredRows += tSubProcessRecord.requiredRows;
        }
        
        var oCounterpartEventOccList = oCurrentEventOcc.ObjDef().OccList();
        for (countCounterpartEvents = 0; countCounterpartEvents < oCounterpartEventOccList.length; countCounterpartEvents++)
        {
            var oCurrentCounterpartEventOcc = oCounterpartEventOccList[countCounterpartEvents];
            // Is the assumed counterpart event occ not a counterpart event occ but this occ itself?
            if (oCurrentCounterpartEventOcc.X() == oCurrentEventOcc.X() &&
                oCurrentCounterpartEventOcc.Y() == oCurrentEventOcc.Y() &&
                oCurrentCounterpartEventOcc.Model().GUID() == oCurrentEventOcc.Model().GUID())
                oCurrentCounterpartEventOcc = null;
            // Is the counterpart event occ a start event?
            else if (oCurrentCounterpartEventOcc.Cxns(Constants.EDGES_IN).length == 0)
            {   ;   }
            // Is the counterpart event occ an interface triggered event?
             else if (oCurrentCounterpartEventOcc.Cxns(Constants.EDGES_IN).length == 1 && oCurrentCounterpartEventOcc.Cxns(Constants.EDGES_IN)[0].SourceObjOcc().OrgSymbolNum() == Constants.ST_PRCS_IF)
            {   ;   }
            // Is current event occurency none of this.
            else
                oCurrentCounterpartEventOcc = null;
        
            if (oCurrentCounterpartEventOcc == null)
                continue;
            
            // Prevent multiple processing the same process.
            if (oEvaluatedProcessList.indexOf(oCurrentCounterpartEventOcc.Model().GUID()) >= 0)
                oCurrentCounterpartEventOcc = null;
            else
                oEvaluatedProcessList.push(oCurrentCounterpartEventOcc.Model().GUID());
            
            if (oCurrentCounterpartEventOcc == null)
                continue;
            
            // Check 'end if evaluation' depth.
            if (iCurrentDepth <= iMaxEvaluationDepth)
            {
                tProcessRecord.followupDepth = 1;
                var tSubProcessRecord = evaluateProcess(tProcessRecord, oCurrentCounterpartEventOcc.Model(), oCurrentCounterpartEventOcc, iCurrentDepth + 1);
                tProcessRecord.successors.push(tSubProcessRecord);
                tProcessRecord.followupDepth = (tProcessRecord.followupDepth > tSubProcessRecord.followupDepth + 1 ? tProcessRecord.followupDepth : tSubProcessRecord.followupDepth + 1);
                
                if (tProcessRecord.successors.length == 1)
                    tProcessRecord.requiredRows  = tSubProcessRecord.requiredRows;
                else
                    tProcessRecord.requiredRows += tSubProcessRecord.requiredRows;
            }
        }
    }
    
    return tProcessRecord;
}

/// <summary> Draw one process record to the indicated sheet and trigger drawing of all it' successor process records. </summary>
/// <param value="oWorksheet"> The {HssfWorksheet} worksheet to create the cell on. </param>
/// <param value="oPatriarch"> The {HssfPatriarch} graphic elements root object. </param>
/// <param value="tProcessRecord"> The {__usertype_processrecord} process record to draw. </param>
/// <param value="iRowOffset"> The {int} offset to the first usable zero-based row index (the cells above the offset are already in use or reserved). </param>
/// <param value="iColOffset"> The {int} offset to the first usable zero-based column index (the cells left the offset are already in use or reserved). </param>
/// <param value="iFktPictureIdx"> The {int} picture reference to use for the function symbol. </param>
/// <param value="iEvtPictureIdx"> The {int} picture reference to use for the event symbol. </param>
/// <param value="iIFaPictureIdx"> The {int} picture reference to use for the interface symbol. </param>
/// <returns> - </returns>
function printoutProcess(oWorksheet, oPatriarch, tProcessRecord, iRowOffset, iColOffset, iFktPictureIdx, iEvtPictureIdx, iIFaPictureIdx)
{
    var iProcessRow = parseInt(((tProcessRecord.requiredRows - 1) / 2 + iRowOffset).toString());
    var iProcessCol = iColOffset;
    var sProcessName = "";
    
    if (tProcessRecord.process != null)
        sProcessName = getFailSaveAttributeValueEx(tProcessRecord.process, Constants.AT_NAME, nLocale, true, false);
    else
        sProcessName = getFailSaveAttributeValueEx(tProcessRecord.deadend, Constants.AT_NAME, nLocale, true, false);
    
    var oHssfProcessRow = oWorksheet.getRow(iProcessRow);
    createCell(oWorksheet, oHssfProcessRow, iProcessCol);
    if (tProcessRecord.process != null)
        drawSymbol(oWorksheet, oPatriarch, iProcessRow, iProcessCol, sProcessName, iFktPictureIdx);
    else
        drawSymbol(oWorksheet, oPatriarch, iProcessRow, iProcessCol, sProcessName, iIFaPictureIdx);
    if (tProcessRecord.predecessor != null && iProcessCol > 0)
        drawConnection(oPatriarch, iProcessCol - 1, iProcessRow, iProcessCol, iProcessRow);
        
    
    var iSubRowOffset = iRowOffset;
    for (var countSubRecords = 0; countSubRecords < tProcessRecord.successors.length; countSubRecords++)
    {
        var tSubProcessRecord = tProcessRecord.successors[countSubRecords];
        var iEventRow  = parseInt(((tSubProcessRecord.requiredRows - 1) / 2 + iSubRowOffset).toString());
        var sEventName = getFailSaveAttributeValueEx(tSubProcessRecord.startEvent.ObjDef(), Constants.AT_NAME, nLocale, true, false); 
        
        var oHssfEventRow = oWorksheet.getRow(iEventRow);
        createCell(oWorksheet, oHssfEventRow, iProcessCol + 1);
        drawSymbol(oWorksheet, oPatriarch, iEventRow, iProcessCol + 1, sEventName, iEvtPictureIdx);
        drawConnection(oPatriarch, iProcessCol, iProcessRow, iProcessCol + 1, iEventRow);
        
        printoutProcess(oWorksheet, oPatriarch, tSubProcessRecord, iSubRowOffset, iProcessCol + 2, iFktPictureIdx, iEvtPictureIdx, iIFaPictureIdx);

        iSubRowOffset =  parseInt((tSubProcessRecord.requiredRows + iSubRowOffset).toString());
    }
}

/// <summary> Create and style a cell. </summary>
/// <param value="oWorksheet"> The {HssfWorksheet} worksheet to create the cell on. </param>
/// <param value="oRow"> The {HssfRow} row to create a cell on. </param>
/// <param value="iColumn"> The {int} zero-based column index to create a cell for. </param>
/// <global value="oHssfCellStyle"> The {HssfCellStyle} default cell style to apply. </global>
/// <returns> The newly created cell. </returns>
function createCell(oWorksheet, oRow, iColumn)
{
    var cell = oRow.createCell(iColumn);
    cell.setCellStyle(oHssfCellStyle);
    oWorksheet.setColumnWidth(iColumn, 8 * 1024);
    return cell;
}

/// <summary> Draw a symbol to it's preferred position into the indicated cell. </summary>
/// <param value="oWorksheet"> The {HssfWorksheet} worksheet to draw on. </param>
/// <param value="oPatriarch"> The {HssfPatriarch} graphic elements root object. </param>
/// <param value="iColNum"> The {int} zero-based column index to use for drawing. </param>
/// <param value="iRowNum"> The {int} zero-based row index to use for drawing. </param>
/// <param value="sText"> The {string} text to add to the symbol. </param>
/// <param value="iPictureIdx"> The {int} picture reference to use for the symbol to draw. </param>
/// <returns> - </returns>
function drawSymbol(oWorksheet, oPatriarch, iRowNum, iColNum, sText, iPictureIdx)
{
    var iCorrectionOffset = 4;
    var iSymbolHeight = 32;
    
    var oHssfRow = oWorksheet.getRow(iRowNum);
    var oHssfCell = oHssfRow.getCell(iColNum);
    oHssfCell.setCellValue("\n" + sText);
    var oHssfAnchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(5 * 1024 / 12, iCorrectionOffset, 7 * 1024 / 12, iCorrectionOffset + iSymbolHeight, iColNum, iRowNum, iColNum, iRowNum );
    //var oHssfShape = oPatriarch.createSimpleShape(oHssfAnchor);
    //oHssfShape.setShapeType(org.apache.poi.hssf.usermodel.HSSFSimpleShape.OBJECT_TYPE_RECTANGLE); // OBJECT_TYPE_OVAL
    //oHssfShape.setLineStyleColor(64,196,64);                                                      // (196,64,128)
    //oHssfShape.setFillColor(128,255,128);                                                         // (255,128,196)
    var oHssfPicture = oPatriarch.createPicture(oHssfAnchor, iPictureIdx);
}

/// <summary> Draw a connection from a start symbol's preferred position to an end symbol's preferred position. </summary>
/// <param value="oPatriarch"> The {HssfPatriarch} graphic elements root object. </param>
/// <param value="iStartCol"> The {int} zero-based start symbol's column index. </param>
/// <param value="iStartRow"> The {int} zero-based start symbol's row index. </param>
/// <param value="iEndCol"> The {int} zero-based end symbol's column index. </param>
/// <param value="iEndRow"> The {int} zero-based end symbol's row index. </param>
/// <returns> - </returns>
function drawConnection(oPatriarch,  iStartCol, iStartRow, iEndCol, iEndRow)
{
    var iCorrectionOffset = 4;
    var iSymbolHeight = 32;
    var iArrowLine1 = 8;
    var iArrowLine2 = 5;
    var iArrowLine3 = 2;
    
    var oHssfAnchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(7 * 1024 / 12 - iCorrectionOffset,     iCorrectionOffset + iSymbolHeight / 2 - iCorrectionOffset / 2,
                                                                         5 * 1024 / 12 - iCorrectionOffset * 2, iCorrectionOffset + iSymbolHeight / 2 - iCorrectionOffset / 2, iStartCol, iStartRow, iEndCol, iEndRow);
    var oHssfShape  = oPatriarch.createSimpleShape(oHssfAnchor);
    oHssfShape.setShapeType(org.apache.poi.hssf.usermodel.HSSFSimpleShape.OBJECT_TYPE_LINE);
    oHssfShape.setLineStyleColor(0,0,0);
    oHssfShape.setFillColor(0,0,0);

    oHssfAnchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(5 * 1024 / 12 - 16, iCorrectionOffset + iSymbolHeight / 2 - iArrowLine1 - iCorrectionOffset / 2,
                                                                     5 * 1024 / 12 - 16, iCorrectionOffset + iSymbolHeight / 2 + iArrowLine1 - iCorrectionOffset / 2, iEndCol, iEndRow, iEndCol, iEndRow);
    var oHssfShape  = oPatriarch.createSimpleShape(oHssfAnchor);
    oHssfShape.setShapeType(org.apache.poi.hssf.usermodel.HSSFSimpleShape.OBJECT_TYPE_OVAL);
    oHssfShape.setLineStyleColor(0,0,0);
    oHssfShape.setFillColor(0,0,0);
    oHssfShape.setLineWidth(org.apache.poi.hssf.usermodel.HSSFShape.LINEWIDTH_ONE_PT * 1);

    oHssfAnchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(5 * 1024 / 12 - 12, iCorrectionOffset + iSymbolHeight / 2 - iArrowLine2 - iCorrectionOffset / 2,
                                                                     5 * 1024 / 12 - 12, iCorrectionOffset + iSymbolHeight / 2 + iArrowLine2 - iCorrectionOffset / 2, iEndCol, iEndRow, iEndCol, iEndRow);
    var oHssfShape  = oPatriarch.createSimpleShape(oHssfAnchor);
    oHssfShape.setShapeType(org.apache.poi.hssf.usermodel.HSSFSimpleShape.OBJECT_TYPE_OVAL);
    oHssfShape.setLineStyleColor(0,0,0);
    oHssfShape.setFillColor(0,0,0);
    oHssfShape.setLineWidth(org.apache.poi.hssf.usermodel.HSSFShape.LINEWIDTH_ONE_PT * 1);

    oHssfAnchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(5 * 1024 / 12 - 6, iCorrectionOffset + iSymbolHeight / 2 - iArrowLine3 - iCorrectionOffset / 2,
                                                                     5 * 1024 / 12 - 6, iCorrectionOffset + iSymbolHeight / 2 + iArrowLine3 - iCorrectionOffset / 2, iEndCol, iEndRow, iEndCol, iEndRow);
    var oHssfShape  = oPatriarch.createSimpleShape(oHssfAnchor);
    oHssfShape.setShapeType(org.apache.poi.hssf.usermodel.HSSFSimpleShape.OBJECT_TYPE_OVAL);
    oHssfShape.setLineStyleColor(0,0,0);
    oHssfShape.setFillColor(0,0,0);
    oHssfShape.setLineWidth(org.apache.poi.hssf.usermodel.HSSFShape.LINEWIDTH_ONE_PT * 1);

    oHssfAnchor = new org.apache.poi.hssf.usermodel.HSSFClientAnchor(5 * 1024 / 12 - 16, iCorrectionOffset + iSymbolHeight / 2 - iCorrectionOffset / 2,
                                                                     5 * 1024 / 12 - 4,  iCorrectionOffset + iSymbolHeight / 2 - iCorrectionOffset / 2, iEndCol, iEndRow, iEndCol, iEndRow);
    var oHssfShape  = oPatriarch.createSimpleShape(oHssfAnchor);
    oHssfShape.setShapeType(org.apache.poi.hssf.usermodel.HSSFSimpleShape.OBJECT_TYPE_OVAL);
    oHssfShape.setLineStyleColor(0,0,0);
    oHssfShape.setFillColor(0,0,0);
    oHssfShape.setLineWidth(org.apache.poi.hssf.usermodel.HSSFShape.LINEWIDTH_ONE_PT * 1);
}

/// <summary> Get an item's indicated attribute value. Avoid access to a null item. </summary>
/// <param value="oItem"> The {ObjDef, Model or Group} item to get the attribute value from. </param>
/// <param value="iAttrNum"> The {int} number of the attribute value to get. </param>
/// <param value="iLanguage"> The {int} language of the attribute value to get.. </param>
/// <param value="bUseFallbackLanguage"> The {bool} determineation whether to use fallback language, if requested language is not maintained for indicated attribute. </param>
/// <param value="bRemoveLineBreak"> The {bool} determineation whether to remove linebreaks. </param>
/// <returns> The attribute value on success, or an empty string on any error. </returns>
function getFailSaveAttributeValueEx(oItem, iAttrNum, iLanguage, bUseFallbackLanguage, bRemoveLineBreak)
{
    if (oItem == null)
        return "(unknown)";

    var oAttr = oItem.Attribute(iAttrNum, iLanguage);
    if (!oAttr.IsMaintained() && bUseFallbackLanguage)
    {
        oAttr = oItem.Attribute(iAttrNum, 1031);
        if (!oAttr.IsMaintained())
            return "(unknown)";
        else
            return "(" + getFailSaveAttributeValue(oAttr, bRemoveLineBreak) + ")";
    }
    return getFailSaveAttributeValue(oAttr, bRemoveLineBreak);
}

/// <summary> Get an item's indicated attribute value. Avoid access to not maintained value. </summary>
/// <param value="oAttr"> The {AttrDef} attribute to get the value from. </param>
// <param value="bRemoveLineBreak"> The {bool} determineation whether to remove linebreaks. </param>
/// <returns> The attribute value on success, or an empty string on any error. </returns>
function getFailSaveAttributeValue(oAttr, bRemoveLineBreak)
{
    if (oAttr.IsMaintained() == true)
    {
        var result = oAttr.GetValue(bRemoveLineBreak);
        if (result == null)
            return "";
        else
            return result;
    }
    else
        return "";
}
