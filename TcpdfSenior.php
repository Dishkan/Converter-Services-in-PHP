<?php

namespace zetsoft\service\App\eyuf;
namespace zetsoft\service\office;

use zetsoft\system\kernels\ZFrame;

class Tcpdf extends ZFrame
{
    public function test($path)
    {
        require_once  __DIR__.'\..\..\vendor\autoload.php';



        $objReader = \PhpOffice\PhpWord\IOFactory::createReader('Word2007');
        $contents = $objReader->load($path);

        $rendername = \PhpOffice\PhpWord\Settings::PDF_RENDERER_TCPDF;
        chdir('D:\Develop\Projects\ALL\asrorz\zetsoft\vendor\tecnickcom');
        $renderLibrary="TCPDF";
        $renderLibraryPath=''.$renderLibrary;
        if (!\PhpOffice\PhpWord\Settings::setPdfRenderer($rendername, $renderLibrary)) {
            die("Provide Render Library And Path");
        }
        $renderLibraryPath = '' . $renderLibrary;
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($contents, 'PDF');
        $objWriter->save("D:/office.pdf");


    }


    public function setUp()
    {
        //     $this->markTestSkipped(); // skip this test
        $this->obj = new \Com\Tecnick\Pdf\Tcpdf();
    }

    public function testDummy()
    {
        $this->assertEquals(1, 1);
    }
}
