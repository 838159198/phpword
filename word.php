<?php
// echo "<title>导出word</title>";
// echo "<meta charset='utf-8'>";
require_once"PHPWord.php";
$PHPWord = new PHPWord();
			//1,文档默认字体、字号
            $PHPWord->setDefaultFontName('宋体');
            $PHPWord->setDefaultFontSize(10.5);
            //文本样式
            $PHPWord->addFontStyle('rStyle', array(
                'bold' => true,
                'italic' => true,
                'size' => 10.5,
            ));
            $fontstyle=array(
            	'bold' => true,
                'italic' => true,
                'size' => 10.5,
            	);
            $PHPWord->addFontStyle('tStyle', array(
                'bold' => true,
                'italic' => false,
                'size' => 10.5,
            ));
            $PHPWord->addParagraphStyle('pStyle', array(
                'align' => 'center',
                'spaceAfter' => 100,
            ));
            $PHPWord->addTitleStyle('title', array(
                'bold' => true,
                'size' => 14,
            ), array(
                'spaceAfter' => 100,
            ));
            //添加页面
            $section = $PHPWord->createSection();
            //添加标题
            $section->addTitle('大大大','title');
            //添加文本
            $section->addText('啊嗒嗒嗒嗒');
            //添加文本资源使用函数方法createTextrun.
            $textrun = $section->createTextRun('rStyle');//调用文本样式pStyle，rStyle
            $textrun->addText('我是一个',$fontstyle);
            //添加超链接
            $section->addLink('http://www.baidu.com','草链接',$fontstyle);
            //添加图片
            $imagestyle=array('width'=>40,'height'=>40);
            $section->addImage('aaa.jpg',$imagestyle); 
            //addTOC添加目录
            $section->addTOC();
            //添加文档页脚使用函数方法： createFooter:
            $footer = $section->createFooter(); 
            $footer->addText('我是页脚');
            // $footer->addPreserveText('Page {PAGE} of {NUMPAGES}.'); //添加页码使用函数方法：addPreserveText: 
            //添加换行符
            $section->addTextBreak(); 
            header('Content-Disposition: attachment;filename="paper.docx"'); //下载 doc
            header("Content-Type: application/octet-stream");
            // IE 与 ssl 需要的信息
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
            header('Cache-Control: cache, must-revalidate');
            header('Pragma: public');

            $objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
            $objWriter->save('php://output');






?>