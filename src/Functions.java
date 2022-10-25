import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import javax.swing.JOptionPane;

import static java.nio.file.StandardCopyOption.*;
public class Functions 
{
    List<String> fileName = new ArrayList<String>();
    List<String> folderList = new ArrayList<String>();
    String[] archivos1;
    String[] ficheros;
    String nuevoParent;
    File zip;
    ZipOutputStream output;

    
    //Función que recorre la carpeta para encontrar los archivos .xml y hacer la conversión del string
    public void findXMLfiles(File folder) 
    {
        for (File file : folder.listFiles()) 
        {
			if (file.isDirectory()) 
            {
                if(file.getName().equalsIgnoreCase("worksheets"))
                {
                    System.out.println("Encontrada la carpeta "+file.getName());
                    System.out.println("Recuperando las hojas del libro");
                    File[] hojas = file.listFiles();
			        System.out.println("Número de hojas encontradas: "+hojas.length);
                    System.out.println("Iniciando desbloqueo de hojas encontradas");
                    for(int i=1;i<=hojas.length; i++)
                    {
                        try
                        {
                            System.out.println("Abriendo archivo sheet"+i+".xml");
                            InputStream ins = new FileInputStream(file.getPath().toString()+"\\sheet"+i+".xml");
                            Scanner obj = new Scanner(ins);
                            String palabra;
                            FileWriter escribir;
                            while (obj.hasNextLine())
                            {
                                palabra=obj.nextLine();
                                palabra=searchWordstoreplace(palabra, "sheet=\"1\"", "sheet=\"0\"");
                                palabra=searchWordstoreplace(palabra, "objects=\"1\"", "objects=\"0\"");
                                palabra=searchWordstoreplace(palabra, "scenarios=\"1\"", "scenarios=\"0\"");
                                //Creando archivos .XML con las modificaciones para su desbloqueo
                                File archivo = new File(file.getPath().toString()+"\\sheet"+i+".xml");
                                escribir = new FileWriter(archivo);
                                escribir.write(palabra);
                                escribir.close();
                            } 
                            System.out.println("Fin de lectura para el archivo sheet"+i+".xml");
                            obj.close();
                        }
                        catch(Exception e)
                        {
                            System.err.println(e);
                        }
                    }
                }
                else
                {
                    findXMLfiles(file);
                }  
			} 
		}
    }

    public void copyAndmovefile(File originalFile,File copyFile, File saveFilelocation)
    {
        try
        {
            Files.copy(originalFile.toPath(), copyFile.toPath(),REPLACE_EXISTING); 
            File newFolder = new File (saveFilelocation.getPath().toString().replaceAll(saveFilelocation.getName(), saveFilelocation.getName().toString().replaceAll(".xlsx"," (Desbloqueado)")));
            if (!newFolder.exists()) 
            {
                newFolder.mkdir();
                copyFile.renameTo(new File(newFolder.toPath()+"\\"+copyFile.getName()));
                System.out.println("Directory created successfully");
            }
            String directorioZip = newFolder.toPath().toString()+"\\"+copyFile.getName().toString();
            unzipFile(directorioZip,newFolder);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    public void unzipFile(String directorioZip, File  newFolder)
    {
        try
        {
            File carpetaExtraer = new File(directorioZip);
		//valida si existe el directorio
		    if (carpetaExtraer.exists()) 
            {
			    System.out.println("Descomprimiendo.....");
                byte[] buffer = new byte[1024];
                try
                {
                    ZipInputStream zis = new ZipInputStream(new FileInputStream(directorioZip));
                    ZipEntry ze = zis.getNextEntry();
                    while (ze != null) 
                    {
                        String nombreArchivo = ze.getName();
                        File archivoNuevo = new File(newFolder.getPath()+ File.separator + nombreArchivo);
                        System.out.println("archivo descomprimido : " + archivoNuevo.getAbsoluteFile());
                        new File(archivoNuevo.getParent()).mkdirs();
                        FileOutputStream fos = new FileOutputStream(archivoNuevo);
                        int len;
                        while ((len = zis.read(buffer)) > 0) 
                        {
                            fos.write(buffer, 0, len);
                        }
                        fos.close();
                        ze = zis.getNextEntry();
                    }
                    zis.closeEntry();
                    zis.close();
                    carpetaExtraer.delete();
                    System.out.println("Listo");
                    System.out.println("Buscando hoja para desbloquear");
                    File newFolder2 = new File (newFolder.toPath().toString()+"\\");
                    findXMLfiles(newFolder2);
                    newFolder2.setWritable(true);
                    newFolder2.setReadable(true);
                
                    System.out.println("Iniciando conversión final...");
                    //String parent = carpetaExtraer.getParent();
                    String parent = carpetaExtraer.getParent();
                    //String nuevoParent = parent.replaceAll("\\\\", "\\\\\\\\")+"\\";
                    nuevoParent = parent;
                    JOptionPane.showMessageDialog(null, "CARPETA SELECCIONADA -> " + nuevoParent);
                    String destino = carpetaExtraer.getPath().toString();
                    JOptionPane.showMessageDialog(null, "DESTINO -> " + destino);
                    JOptionPane.showMessageDialog(null, "Comprimiendo...");
                    output = new ZipOutputStream(new FileOutputStream(destino));
                    recorrer(newFolder2);
                }
                catch (Exception e)
                {
                    e.printStackTrace();
                }
            }
        }
        catch (Exception e)
        {
            e.printStackTrace();
       }
    }

  
    //Función para buscar y reemplazar valores en el string para el desbloqueo del archivo
    public String searchWordstoreplace(String vPalabra, String cambio, String sFinal)
    {
        if(vPalabra.indexOf(cambio)!=-1)
        {
            System.out.println("Encontrada la frase sheet");
            System.out.println("Reemplazando el valor de sheet...");
            vPalabra=vPalabra.replaceAll(cambio, sFinal);
        }
        return vPalabra;
    }
    
    public void recorrer(File recorrido)
    {
        for (File f : recorrido.listFiles()) 
        {
            if (f.isFile()) 
            {
                //System.out.println("File: " + f.getName());
                zipFile(recorrido);
            }
            if (f.isDirectory()) 
            {
                //System.out.println("Carpeta: " + f.getName());
                zipDir(recorrido);
                recorrer(f);
            }
        }
    }
    private void zipDir(File file) 
    {
        try 
        {
            System.out.println(file.getPath()+"\\");
            output.putNextEntry(new ZipEntry(file.getPath()+"\\"));
            output.closeEntry();
        
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }
    private void zipFile(File file)
    {
        try 
        {
            byte[] buf = new byte[1024];
            output.putNextEntry(new ZipEntry(file.getPath()));
            FileInputStream fis = new FileInputStream(file);
              int len;
              while ((len = fis.read(buf)) > 0) {
                output.write(buf, 0, len);
              }
              fis.close();
              output.closeEntry();
              
        } 
        catch (IOException e) 
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
    }

  

    //  public void agregarCarpeta(String ruta, String carpeta, ZipOutputStream zip) throws Exception {
        
    //     if(!carpeta.toString().equals(nuevoParent))
    //     {
    //         System.out.println("Ruta: "+ruta);
    //        System.out.println("Carpeta: "+carpeta);
    //         File directorio = new File(carpeta);
    //         for (String nombreArchivo : directorio.list()) 
    //         {
    //             if (ruta.equals("")) 
    //             {
    //                 /* JOptionPane.showMessageDialog(null, "ruta vale: "+ruta);
    //                 JOptionPane.showMessageDialog(null, "nombre del directorio: "+directorio.getName()); */
    //                 agregarArchivo(directorio.getName(), carpeta + "/" + nombreArchivo, zip);
    //             }
    //              else 
    //             {
    //                 // JOptionPane.showMessageDialog(null, "ruta equivale a:" +ruta);
    //                 agregarArchivo(ruta + "/" + directorio.getName(), carpeta + "/" + nombreArchivo, zip);
    //             }
    //         }
    //     }
    //     else
    //     {
           
    //        System.out.println("Ruta1: "+ruta);
    //        System.out.println("Carpeta1: "+carpeta);
    //        agregarArchivo("", carpeta+ "/", zip);
    //     }
    // }
    // public void agregarArchivo(String ruta, String directorio, ZipOutputStream zip) throws Exception 
    // {
        
    //     //JOptionPane.showMessageDialog(null, "valor de directorio: "+directorio);
       
      
    //     File archivo = new File(directorio);
    //     if (archivo.isDirectory()) 
    //     {
    //         // JOptionPane.showMessageDialog(null, "El valor es de un directorio");
    //         agregarCarpeta(ruta, directorio, zip);
    //     } 
    //     else 
    //     {
    //         // JOptionPane.showMessageDialog(null, "El valor es de un archivo");
    //         if(!archivo.getName().toString().equals("notas laboral (Desbloqueado).zip"))
    //           {
    //          byte[] buffer = new byte[4096];
    //             int leido;
    //             FileInputStream entrada = new FileInputStream(archivo);
    //             zip.putNextEntry(new ZipEntry(ruta + "/" + archivo.getName()));
    //             while ((leido = entrada.read(buffer)) > 0) 
    //             {
    //                 zip.write(buffer, 0, leido);
    //             }
    //         }
    //     }
    // } 

}


