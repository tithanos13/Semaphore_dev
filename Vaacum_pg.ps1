$SmtpServer = "smtp-relay.gmail.com"
$SMTPport = 587
$encodingMail = "UTF8"
$MailSender = "exploitation.info@alinea.com"
$ReceiverMail = "romain.valdebon@alinea.com", "jeremy.carpentier@domaine.com"

$SubjectFailure = "Rapport Vacuum - Echec"
$BodyFailure = @"
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
<style="color:rgb(0,0,0)"><font face="verdana, sans-serif" style="color:rgb(0,0,0)"><font color="#274e13">
<br>Bonjour,</br>
<br> </br>
<br>Echec du Vacuum</br>
<br> </br>
</body>
</html>
"@

try {
    # 1. Arrêt des services Horoquartz
    Write-Host "Arret des services Horoquartz..."
    Invoke-Command -ComputerName "SRV-HORO-DEV" -ScriptBlock { etemptation stop all }

    # 2. Exécution du Vacuum et mesure du temps
    Write-Host "Execution de la commande Vacuum (psql)..."
    $executionTime = Measure-Command {
        psql -U postgres -d etemptation -c "VACUUM FULL VERBOSE ANALYSE;"
    }

    # 3. Récupération de la date et de l'heure actuelles
    $executionDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # 4. Création du contenu du mail de succès
    $BodySuccess = @"
<!DOCTYPE html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
</head>
<body>
<style="color:rgb(0,0,0)"><font face="verdana, sans-serif" style="color:rgb(0,0,0)"><font color="#274e13">
<br>Bonjour,</br>
<br> </br>
<br>Vacuum effectué avec succès.</br>
<br> </br>
<p>Date d'exécution : **$executionDate**</p>
<p>Détails de l'exécution :</p>
<table border="1" cellpadding="5" cellspacing="0" style="font-family:verdana, sans-serif; font-size:12px;">
    <tr><td>Days</td><td>$($executionTime.Days)</td></tr>
    <tr><td>Hours</td><td>$($executionTime.Hours)</td></tr>
    <tr><td>Minutes</td><td>$($executionTime.Minutes)</td></tr>
    <tr><td>Seconds</td><td>$($executionTime.Seconds)</td></tr>
    <tr><td>Milliseconds</td><td>$($executionTime.Milliseconds)</td></tr>
    <tr><td>Ticks</td><td>$($executionTime.Ticks)</td></tr>
    <tr><td>TotalDays</td><td>$($executionTime.TotalDays)</td></tr>
    <tr><td>TotalHours</td><td>$($executionTime.TotalHours)</td></tr>
    <tr><td>TotalMinutes</td><td>$($executionTime.TotalMinutes)</td></tr>
    <tr><td>TotalSeconds</td><td>$($executionTime.TotalSeconds)</td></tr>
    <tr><td>TotalMilliseconds</td><td>$($executionTime.TotalMilliseconds)</td></tr>
</table>
<br> </br>
</body>
</html>
"@

    # 5. Envoi du mail de succès
    Write-Host "Envoi du mail de succes..."
    Send-MailMessage -To $ReceiverMail `
                     -From $MailSender `
                     -Subject "Rapport Vacuum - Success" `
                     -SmtpServer $SMTPserver `
                     -BodyAsHtml $BodySuccess `
                     -Port $SMTPport `
                     -Encoding $encodingMail `
                     -Priority High

} catch {
    # Si le bloc try a échoué, cette partie s'exécute pour envoyer l'e-mail d'échec
    Write-Host "Une erreur est survenue. Envoi du mail d'echec..." -ForegroundColor Red
    Send-MailMessage -To $ReceiverMail `
                     -From $MailSender `
                     -Subject $SubjectFailure `
                     -SmtpServer $SMTPserver `
                     -BodyAsHtml $BodyFailure `
                     -Port $SMTPport `
                     -Encoding $encodingMail `
                     -Priority High
} finally {
    # Le bloc finally s'exécute qu'il y ait une erreur ou non
    Write-Host "Demarrage des services Horoquartz..."
    Invoke-Command -ComputerName "SRV-HORO-DEV" -ScriptBlock { etemptation start all }
}