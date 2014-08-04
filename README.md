Excel Post Json Add In
======================

Excel plugin to post any table to remote URL

https://github.com/mac2000/ExcelPostJsonAddIn/raw/gh-pages/publish.zip - binaries

How to use
----------

[![Excel Post Json Add In](http://img.youtube.com/vi/cENfLmC6dlc/0.jpg)](http://www.youtube.com/watch?v=cENfLmC6dlc)

Backend
-------

Here is simple example used in video:

    <?php

    if(@$_SERVER['PHP_AUTH_USER'] != 'admin' || @$_SERVER['PHP_AUTH_PW'] != '123') {
        header('WWW-Authenticate: Basic realm="Post Json"');
        header('HTTP/1.0 401 Unauthorized');
        exit;
    }

    echo 'REQUEST_METHOD: ' . $_SERVER['REQUEST_METHOD'] . PHP_EOL . PHP_EOL;

    echo 'REQUEST:' . PHP_EOL;
    print_r($_REQUEST);
    echo PHP_EOL . PHP_EOL;

    echo 'INPUT:' . PHP_EOL;
    print_r(json_decode(file_get_contents('php://input'), true));
    echo PHP_EOL . PHP_EOL;

    echo 'PHP_AUTH_USER: ' . @$_SERVER['PHP_AUTH_USER'] . PHP_EOL;
    echo 'PHP_AUTH_PW: ' . @$_SERVER['PHP_AUTH_PW'] . PHP_EOL . PHP_EOL;
