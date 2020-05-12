<?php
    ////////////////////////////////////////////////
    // PHP Office VBA password remover.
    //
    // Created by Kally for /r/excel
    // Feel free to do whatever you want
    // as long as access to this stays free of
    // any charges and/or subscriptions.
    //
    // Working copy at 
    //
    // You need the bootstrap framework for this to look nice.
    // It's not required for the functionality itself.
    //
    ///////////////////////////////////////////////

    //Be extra pedantic
    error_reporting(E_ALL);

    //Names of VBA blobs to search for in zipped office formats
    define('NAMES','xl/vbaProject.bin|word/vbaProject.bin|ppt/vbaProject.bin');

    //Holds error message for the user
    $err='';

    //Extract vbaProject.bin from a modern office file.
    function getFromZip($fName){
        global $err;
        $temp='';
        $zip=new ZipArchive();

        //Try to open the file as zip archive
        if($res=$zip->open($fName) && $zip->numFiles>0){
            foreach(explode('|',NAMES) as $name){
                if($temp=$zip->getFromName($name)){
                    $zip->close();
                    //return the vbaProject.bin content if it has been extracted.
                    return $temp;
                }
            }
            $zip->close();
        }
        else{
            $err='Can\'t open your file. It seems to be from Office 2007 or newer but is corrupt.';
        }
        //No matching file found
        return '';
    }

    //Add vbaProject.bin back to a modern Office file
    function addToZip($contents,$fName)
    {
        global $err;
        $temp='';
        $zip=new ZipArchive;
        //Open file as zip archive
        if($res=$zip->open($fName))
        {
            //Try to find where the original vbaProject.bin was located and replace it with the unlocked content.
            foreach(explode('|',NAMES) as $name){
                if($zip->getFromName($name)){
                    $zip->deleteName($name);
                    $zip->addFromString($name,$contents);
                    $zip->close();
                    return TRUE;
                }
            }
            $zip->close();
        }
        else{
            $err='Can\'t open file to change VBA settings. Inform the administrator.';
        }
        return FALSE;
    }

    //Convert file name into safe name by replacing problematic characters
    function safeName($x){
        //Replaces:
        //- Control characters
        //- Invalid characters \/"?*<>|
        //- High ASCII characters
        return preg_replace('#[\x00-\x1F"<>\|\?\*\\/:\x7F-\xFF]#','_',$x);
    }

    //provides file source code upon request
    if(isset($_GET['source'])){
        echo '<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>VBA Unlocker Source</title>
</head>
<body>';
        highlight_file(__FILE__);
        echo '</body></html>';
        exit(0);
    }

    //Check if file uploaded
    if(isset($_FILES['excel'])){
        //Ensure a file name exists
        if(isset($_FILES['excel']['tmp_name']) && $_FILES['excel']['tmp_name']!=''){
            //Try to read the file
            if($fp=fopen($_FILES['excel']['tmp_name'],'rb')){
                //Read everything and close file handle
                //Note: No need to delete the uploaded file. PHP does this for us
                $contents=fread($fp,filesize($_FILES['excel']['tmp_name']));
                fclose($fp);

                //If it starts with 'PK' it (probably) is a modern file (O2007 and newer)
                if(substr($contents,0,2)=='PK'){
                    //Create temporary zip file name
                    $z=tempnam(dirname(__FILE__).'/TMP/','zip');
                    //Move temporary file to place where it's not readonly to us.
                    move_uploaded_file($_FILES['excel']['tmp_name'], $z);
                    //Get VBA blob from zip
                    $contents=getFromZip($z);
                    if($contents!='' && $err==''){
                        if(strpos($contents,'DPB=')===FALSE){
                            $err='We found VBA Code but it is not protected.';
                        }
                        else{
                            $contents=str_replace('DPB=','DPx=',$contents);
                            addToZip($contents,$z);
                            if($err=='')
                            {
                                if($fp=fopen($z,'rb')){
                                    $length=filesize($z);
                                    header('Content-Type: application/octet-stream');
                                    header("Content-Length: $length");
                                    header('Content-Disposition: attachment; filename="' . safeName($_FILES['excel']['name']) . '"');
                                    echo fread($fp,$length);
                                    fclose($fp);
                                    //Delete the uploaded file on success
                                    @unlink($z);
                                    exit(0);
                                }
                                else{
                                    $err='Can\'t send back office file. Error opening the temporary file.';
                                }
                            }
                        }
                    }
                    else{
                        $err='This file has no encrypted VBA code, or the entire file is encrypted.';
                    }
                    //Delete the uploaded file on error
                    @unlink($z);
                }
                else{
                    //Delete uploaded file because it's in $contents now
                    unlink($_FILES['excel']['tmp_name']);

                    //assume classic file (O2003 and older)
                    if(strpos($contents,'DPB=')===FALSE){
                        $err='There is no VBA code or it is not protected.';
                    }
                    else{
                        //This removes the protection
                        $contents=str_replace('DPB=','DPx=',$contents);
                        //Send back file
                        header('Content-Length: ' . strlen($contents));
                        header('Content-Disposition: attachment; filename="' . safeName($_FILES['excel']['name']) . '"');
                        header('Content-Type: application/octet-stream');
                        echo $contents;
                        exit(0);
                    }
                }
            }
            else{
                $err='We were unable to open the file. Either our disk is full or it was removed by our Anti-Virus';
            }
        }
        else{
            $err='No file received. Please select a file to decrypt';
        }
    }
?><!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=edge" />
        <meta name="viewport" content="width=device-width, initial-scale=1" />
        <link rel="stylesheet" type="text/css" href="/bootstrap4/bootstrap.min.css" />
        <script type="text/javascript" src="/bootstrap4/jquery-3.2.1.min.js"></script>
        <title>Office VBA Password remover</title>
    </head>
    <body>
        <div class="container">
            <h1>Office VBA Password remover</h1>
            <?php if($err){ echo "<div class='alert alert-danger'>$err</div>";} ?>
            <form method="post" action="index.php" enctype="multipart/form-data">
                <label class="control-label">Office File (word, excel, powerpoint):
                <input type="file" name="excel" class="form-control" required accept=".doc,.docm,.xls,.xlsm,.ppt,.pptm"/></label><br />
                <input type="button" class="btn btn-secondary" value="Open file" style="display:none" />
                <input type="submit" class="btn btn-primary" value="Decrypt VBA" />
            </form>
            <script>
                $("[type=button]").show().on("click",function(){
                    $("[type=file]")[0].click();
                });
                $("[type=file]").hide();
            </script>
            <h2>Instructions</h2>
            <ol>
                <li>Upload your Office document</li>
                <li>Confirm download of new document</li>
                <li>
                    Open the downloaded document and press <kbd>ALT + F11</kbd>.
                    Confirm all error messages that might appear.
                </li>
                <li>In the Macro window, <b>do not expand the project</b>, go to Tools &gt; VBA Project Properties</li>
                <li>On the "Protection" Tab, set a password of your choice <b>and leave the checkbox selected</b>.</li>
                <li>Save the document and close the Editor</li>
                <li>Repeat Step 3</li>
                <li>On the "Protection" Tab, clear the checkbox and password fields</li>
                <li>Save document again</li>
                <li>The password is now removed and you can view or change the code as if it was never protected</li>
            </ol>
            <hr />
        </div>
    </body>
</html>
