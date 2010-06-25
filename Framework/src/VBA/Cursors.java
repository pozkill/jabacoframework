package VBA;

import javax.swing.*;
import java.awt.event.*;
import java.awt.*;

public class Cursors {

  public static final int VBA_DEFAULT = 0;
  public static final int VBA_ARROW = 1;
  public static final int VBA_CROSSHAIR = 2;
  public static final int VBA_IBEAM = 3;
  public static final int VBA_ICON_POINTER = 4;
  public static final int VBA_SIZE_POINTER = 5;
  public static final int VBA_SIZE_NESW = 6;
  public static final int VBA_SIZE_NS = 7;
  public static final int VBA_SIZE_NWSE = 8;
  public static final int VBA_SIZE_WE = 9;
  public static final int VBA_UP_ARROW = 10;
  public static final int VBA_HOURGLASS = 11;
  public static final int VBA_NO_DROP = 12;
  public static final int VBA_ARROW_HOURGLASS = 13;
  public static final int VBA_ARROW_QUESTION = 14;
  public static final int VBA_SIZE_ALL = 15;
  public static final int VBA_HAND_CURSOR = 16;

  public Cursors() {
  }

  public Cursor getVBACursor(int number) {
    JComponent cp = new JLabel("");  // only for creating the cursors
    switch(number) {
     case VBA_DEFAULT:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_normal.gif")).getImage(), new Point(0, 0), "VB Default Cursor");
     case VBA_ARROW:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_normal.gif")).getImage(), new Point(0, 0), "VB Arrow Cursor");
     case VBA_CROSSHAIR:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_cross.gif")).getImage(), new Point(15, 15), "VB Crosshair Cursor");
     case VBA_IBEAM:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_ibeam.gif")).getImage(), new Point(15, 15), "VB Text Cursor");
     case VBA_ICON_POINTER:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_icon.gif")).getImage(), new Point(7, 6), "VB Iconpointer");
     case VBA_SIZE_POINTER:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_size.gif")).getImage(), new Point(15, 15), "VB Move Cursor");
     case VBA_SIZE_NESW:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_sizenesw.gif")).getImage(), new Point(15, 15), "VB Resize Northeast/Southwest Cursor");
     case VBA_SIZE_NS:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_sizens.gif")).getImage(), new Point(15, 15), "VB Resize North/South Cursor");
     case VBA_SIZE_NWSE:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_sizenwse.gif")).getImage(), new Point(15, 15), "VB Resize Northwest/Southeast Cursor");
     case VBA_SIZE_WE:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_sizewe.gif")).getImage(), new Point(15, 15), "VB Resize West/East Cursor");
     case VBA_UP_ARROW:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_up.gif")).getImage(), new Point(6, 0), "VB Up Cursor");
     case VBA_HOURGLASS:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_wait.gif")).getImage(), new Point(15, 15), "VB Wait Cursor");
     case VBA_NO_DROP:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_no.gif")).getImage(), new Point(15, 15), "VB No-Drop Cursor");
     case VBA_ARROW_HOURGLASS:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_appstarting.gif")).getImage(), new Point(0, 0), "VB Appstarting Cursor");
     case VBA_ARROW_QUESTION:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_help.gif")).getImage(), new Point(0, 0), "VB Help Cursor");
     case VBA_SIZE_ALL:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_sizeall.gif")).getImage(), new Point(15, 15), "VB Size All Cursor");
     case VBA_HAND_CURSOR:
       return cp.getToolkit().createCustomCursor(new ImageIcon(getClass().getResource("cursors/ocr_hand.gif")).getImage(), new Point(7, 0), "VB Hand Cursor");

     default:
       return cp.getToolkit().createCustomCursor(new ImageIcon("cursors/ocr_normal.gif").getImage(), new Point(0, 0), "VB Default Cursor");
    }
  }
}
