$xml = [xml](get-content test.xml)

$value=$xml.appSettings.solutionName.value

write-host $value