<?php

require dirname(__FILE__).'/../vendor/autoload.php';


$body = <<<EOF
<a href="./dojo">使用方法確認</a>
EOF;

?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <title>Document</title>
</head>
<body>
<?php print $body;?>
<hr>
</body>
</html>