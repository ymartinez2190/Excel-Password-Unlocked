import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;

import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.lang.model.util.ElementScanner14;
import javax.swing.JFileChooser;
import static java.nio.file.StandardCopyOption.*;

import java.nio.file.DirectoryStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.util.Scanner;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;
import java.nio.file.StandardCopyOption;
import java.nio.file.Path;
import java.nio.file.Paths;
/**
 *
 * @author Yessenia I. Martínez para AstraCodePTY
 */
public class Window extends javax.swing.JFrame 
{
    File saveFilename;
    File openFilename;
    Functions callFunction=new Functions();

    public Window() {
        initComponents();
    }

    private void initComponents() {

        jpnUp = new javax.swing.JPanel();
        jlbSelection = new javax.swing.JLabel();
        jtfFileroute = new javax.swing.JTextField();
        jbnSearch = new javax.swing.JButton();
        jpnCenter = new javax.swing.JPanel();
        jblFileroute = new javax.swing.JLabel();
        jbnConvertionroute = new javax.swing.JButton();
        jtfConvertroute = new javax.swing.JTextField();
        jpnDown = new javax.swing.JPanel();
        jbnUnlock = new javax.swing.JButton();
        jbnExit = new javax.swing.JButton();
        jmbMenu = new javax.swing.JMenuBar();
        jmnFile = new javax.swing.JMenu();
        jmiExit = new javax.swing.JMenuItem();
        jmnOption = new javax.swing.JMenu();
        jmiLanguage = new javax.swing.JMenuItem();
        jmnHelp = new javax.swing.JMenu();
        jmiManual = new javax.swing.JMenuItem();
        jmiAbout = new javax.swing.JMenuItem();
        jfcFileselector= new javax.swing.JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivo de MS Excel (.xlsx)","xlsx");
        jfcFileselector.setFileFilter(filter);
        

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Desbloqueador de archivos para Excel - Ver 1.0");

        jlbSelection.setText("Seleccione el archivo de Excel: ");

        jbnSearch.setText("Buscar");
        jbnSearch.setToolTipText("Presione \"buscar\" para abrir el cuadro de diálogo y buscar el archivo a desbloquear");
        jbnSearch.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbnSearchActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jpnUpLayout = new javax.swing.GroupLayout(jpnUp);
        jpnUp.setLayout(jpnUpLayout);
        jpnUpLayout.setHorizontalGroup(
            jpnUpLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpnUpLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jpnUpLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jpnUpLayout.createSequentialGroup()
                        .addComponent(jlbSelection)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jpnUpLayout.createSequentialGroup()
                        .addComponent(jbnSearch)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jtfFileroute)))
                .addContainerGap())
        );
        jpnUpLayout.setVerticalGroup(
            jpnUpLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpnUpLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jlbSelection)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jpnUpLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jtfFileroute, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jbnSearch))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jblFileroute.setText("Seleccione la ruta de guardado: ");

        jbnConvertionroute.setText("Guardar en:");
        jbnConvertionroute.setToolTipText("Presione \"guardar en\" para guardar la ruta donde se generará el archivo desbloqueado.");
        jbnConvertionroute.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbnConvertionrouteActionPerformed(evt);
            }
        });

        jtfConvertroute.setToolTipText("");

        javax.swing.GroupLayout jpnCenterLayout = new javax.swing.GroupLayout(jpnCenter);
        jpnCenter.setLayout(jpnCenterLayout);
        jpnCenterLayout.setHorizontalGroup(
            jpnCenterLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpnCenterLayout.createSequentialGroup()
                .addGroup(jpnCenterLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jpnCenterLayout.createSequentialGroup()
                        .addComponent(jblFileroute)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jpnCenterLayout.createSequentialGroup()
                        .addComponent(jbnConvertionroute, javax.swing.GroupLayout.PREFERRED_SIZE, 91, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jtfConvertroute)))
                .addGap(6, 6, 6))
        );
        jpnCenterLayout.setVerticalGroup(
            jpnCenterLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpnCenterLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jblFileroute)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jpnCenterLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jbnConvertionroute)
                    .addComponent(jtfConvertroute, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jbnUnlock.setText("Desbloquear");
        jbnUnlock.setToolTipText("Presione para iniciar el desbloqueo");
        jbnUnlock.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbnUnlockActionPerformed(evt);
            }
        });

        jbnExit.setText("Salir");
        jbnExit.setToolTipText("Presione para salir de la aplicación.");
        jbnExit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jbnExitActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jpnDownLayout = new javax.swing.GroupLayout(jpnDown);
        jpnDown.setLayout(jpnDownLayout);
        jpnDownLayout.setHorizontalGroup(
            jpnDownLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jpnDownLayout.createSequentialGroup()
                .addContainerGap(355, Short.MAX_VALUE)
                .addComponent(jbnUnlock)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jbnExit)
                .addContainerGap())
        );
        jpnDownLayout.setVerticalGroup(
            jpnDownLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jpnDownLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jpnDownLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jbnUnlock)
                    .addComponent(jbnExit))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jmnFile.setText("Archivo");

        jmiExit.setMnemonic('o');
        jmiExit.setText("Salir");
        jmnFile.add(jmiExit);

        jmbMenu.add(jmnFile);

        jmnOption.setText("Configuración");

        jmiLanguage.setText("Idioma");
        jmnOption.add(jmiLanguage);

        jmbMenu.add(jmnOption);

        jmnHelp.setMnemonic('h');
        jmnHelp.setText("Ayuda");

        jmiManual.setText("Manual de uso");
        jmnHelp.add(jmiManual);

        jmiAbout.setText("Acerca de...");
        jmnHelp.add(jmiAbout);

        jmbMenu.add(jmnHelp);

        setJMenuBar(jmbMenu);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jpnUp, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(12, 12, 12)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jpnCenter, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(jpnDown, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(jpnUp, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jpnCenter, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jpnDown, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        );
        pack();
    }

    private void jbnSearchActionPerformed(java.awt.event.ActionEvent evt) 
    {
        if(jfcFileselector.APPROVE_OPTION==0)
        {
            jfcFileselector.showOpenDialog(this);
            openFilename=jfcFileselector.getSelectedFile();
            jtfFileroute.setText(openFilename.getPath());
        }
        else
        {
            JOptionPane.showMessageDialog(null, "No se ha seleccionado un archivo");
        }
    }

    private void jbnConvertionrouteActionPerformed(java.awt.event.ActionEvent evt) 
    {
        
        
            jfcFileselector.showSaveDialog(this);
            saveFilename=jfcFileselector.getSelectedFile();
            jtfConvertroute.setText(saveFilename.getPath());
        
        
    }

    private void jbnUnlockActionPerformed(java.awt.event.ActionEvent evt) 
    {
        try
        {
            File originalFile = new File(openFilename.getPath());  
            File saveFilelocation = new File(saveFilename.getPath());
            File copyFile = new File (originalFile.getPath().toString().replaceAll(".xlsx"," (Desbloqueado).zip"));
            callFunction.copyAndmovefile(originalFile, copyFile,saveFilelocation);
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    } 
    
    private void jbnExitActionPerformed(java.awt.event.ActionEvent evt) {
        if(JOptionPane.showConfirmDialog(null,"¿Desea salir de la aplicación?","Información",JOptionPane.YES_NO_OPTION)==JOptionPane.YES_OPTION)
        {
            this.dispose();
        }
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Window.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Window.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Window.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Window.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
    }

    private javax.swing.JLabel jblFileroute;
    private javax.swing.JButton jbnConvertionroute;
    private javax.swing.JButton jbnExit;
    private javax.swing.JButton jbnSearch;
    private javax.swing.JButton jbnUnlock;
    private javax.swing.JLabel jlbSelection;
    private javax.swing.JMenuBar jmbMenu;
    private javax.swing.JMenuItem jmiAbout;
    private javax.swing.JMenuItem jmiExit;
    private javax.swing.JMenuItem jmiLanguage;
    private javax.swing.JMenuItem jmiManual;
    private javax.swing.JMenu jmnFile;
    private javax.swing.JMenu jmnHelp;
    private javax.swing.JMenu jmnOption;
    private javax.swing.JPanel jpnCenter;
    private javax.swing.JPanel jpnDown;
    private javax.swing.JPanel jpnUp;
    private javax.swing.JTextField jtfConvertroute;
    private javax.swing.JTextField jtfFileroute;
    private javax.swing.JFileChooser jfcFileselector;
}
