<?php
require __DIR__ . '/vendor/autoload.php';

define('ATTACHMENTS_DIR', './attachments', true);

if (isset($_GET['code'])) {
    echo '<div>';
    echo 'Paste the code in your terminal <br /><br />';
    echo 'Code: ' . $_GET['code'];
    echo '</div>';
    die;
}

if (php_sapi_name() != 'cli') {
    throw new Exception('This application must be run on the command line.');
}
$client = getClient();
$sheetsService = new Google_Service_Sheets($client);
$gmailService = new Google_Service_Gmail($client);

/**
 * Returns an authorized API client.
 * @return Google_Client the authorized client object
 */
function getClient()
{
    $client = new Google_Client();
    $client->setApplicationName('Google Sheets API PHP Quickstart');
    $client->setScopes([Google_Service_Sheets::SPREADSHEETS, Google_Service_Gmail::GMAIL_READONLY]);
    $client->setAuthConfig('credentials.json');
    $client->setAccessType('offline');
    $client->setPrompt('select_account consent');

    // Load previously authorized token from a file, if it exists.
    // The file token.json stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.
    $tokenPath = 'token.json';
    if (file_exists($tokenPath)) {
        $accessToken = json_decode(file_get_contents($tokenPath), true);
        $client->setAccessToken($accessToken);
    }

    // If there is no previous token or it's expired.
    if ($client->isAccessTokenExpired()) {
        // Refresh the token if possible, else fetch a new one.
        if ($client->getRefreshToken()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
        } else {
            // Request authorization from the user.
            $authUrl = $client->createAuthUrl();
            printf("Open the following link in your browser:\n%s\n", $authUrl);
            print 'Enter verification code: ';
            $authCode = trim(fgets(STDIN));

            // Exchange authorization code for an access token.
            $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
            $client->setAccessToken($accessToken);

            // Check to see if there was an error.
            if (array_key_exists('error', $accessToken)) {
                throw new Exception(join(', ', $accessToken));
            }
        }
        // Save the token to a file.
        if (!file_exists(dirname($tokenPath))) {
            mkdir(dirname($tokenPath), 0700, true);
        }
        file_put_contents($tokenPath, json_encode($client->getAccessToken()));
    }
    return $client;
}

function getNameFromNumber($num) {
    $numeric = ($num - 1) % 26;
    $letter = chr(65 + $numeric);
    $num2 = intval(($num - 1) / 26);
    if ($num2 > 0) {
        return getNameFromNumber($num2) . $letter;
    } else {
        return $letter;
    }
}

// Function to delete all files
// and directories
function deleteAll($str) {
    // Check for files
    if (is_file($str)) {
        // If it is file then remove by
        // using unlink function
        return unlink($str);
    }
    // If it is a directory.
    elseif (is_dir($str)) {
        // Get the list of the files in this
        // directory
        $scan = glob(rtrim($str, '/').'/*');
        // Loop through the list of files
        foreach($scan as $index=>$path) {
            // Call recursive function
            deleteAll($path);
        }
        // Remove the directory itself
        return @rmdir($str);
    }
}

function readZip($sheetsService) {
    $zip = new ZipArchive();
    $folder = time();

    $zipFiles = glob(ATTACHMENTS_DIR . '/*.zip');

    foreach ($zipFiles as $zipFile) {
        if ($zip->open($zipFile) === true) {
            $zip->setPassword('sproutinsight');
            $zipDest = './' . $folder;
            $zip->extractTo($zipDest);
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $row = 1;
                $filename = $zip->getNameIndex($i);
                $csvDest = './' . $folder . '/' . basename($filename);
                $csvParts = pathinfo($filename);

                if (($handle = fopen($csvDest, 'r')) !== FALSE) {
                    $values = [];
                    $columnCount = 0;
                    $rowCount = 0;
                    while (($data = fgetcsv($handle, 0, ',')) !== FALSE) {
                        $values[] = $data;
                        if (count($data) > $columnCount) {
                            $columnCount = count($data);
                        }
                        $rowCount++;
                    }

                    $spreadsheet = new Google_Service_Sheets_Spreadsheet([
                        'properties' => [
                            'title' => $csvParts['filename']
                        ]
                    ]);
                    $spreadsheet = $sheetsService->spreadsheets->create($spreadsheet, [
                        'fields' => 'spreadsheetId'
                    ]);

                    $body = new Google_Service_Sheets_ValueRange([
                        'values' => $values
                    ]);

                    $params = [
                      'valueInputOption' => 'RAW'
                    ];

                    $sheetsService->spreadsheets_values->update($spreadsheet->spreadsheetId, 'Sheet1!A1:' . getNameFromNumber($columnCount) . $rowCount, $body, $params);
                    fclose($handle);
                }
            }
            deleteAll($zipDest);
            $zip->close();
            unlink($zipFile);
        }
    }
}

function getGmailAttachments ($gmailService) {
    $optParams['q'] = 'From:james@startsmartsourcing.com';
    $msgConf = $gmailService->users_messages->listUsersMessages('me', $optParams);
    $messages = $msgConf->getMessages();
    $files = [];

    foreach ($messages as $message) {
        $msg = $gmailService->users_messages->get('me', $message->getId());
        $parts = $msg->getPayload()->getParts();
        foreach ($parts as $part) {
            if (!is_null($part['body']->attachmentId)) {
                $attachmentObj = $gmailService->users_messages_attachments->get('me', $message->getId(), $part['body']->attachmentId);
                $rawData = $attachmentObj->getData(); //Get data from attachment object
                $sanitizedData = strtr($rawData,'-_', '+/');
                $decodedMessage = base64_decode($sanitizedData);

                $filename = uniqid() . '.zip';
                if (!is_dir(ATTACHMENTS_DIR)) {
                    mkdir(ATTACHMENTS_DIR, 0777);
                }
                // this will empty the memory and appen your zip content
                file_put_contents(ATTACHMENTS_DIR . '/' . $filename, $decodedMessage);
            }
        }
    }
}
getGmailAttachments($gmailService);
readZip($sheetsService);