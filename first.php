<?php

require_once 'vendor/autoload.php'; //autoload do composer

use PhpOffice\PhpPresentation\PhpPresentation; //classe do PhpPresentation
use PhpOffice\PhpPresentation\IOFactory; //classe para manipular os arquivos
use PhpOffice\PhpPresentation\Style\Color; //classe de estilo de cores
use PhpOffice\PhpPresentation\Style\Alignment; //classe de estilo de alinhamentos

$objPHPPowerPoint = new PhpPresentation(); //instanciando uma nova apresentação

$currentSlide = $objPHPPowerPoint->getActiveSlide(); //retornando o slide ativo

$shape = $currentSlide->createDrawingShape(); //criando uma forma de desenho (imagem)
$shape->setName('PokePHP Logo') //definindo o nome da forma
      ->setDescription('PokePHP Logo') //definindo a descrição da forma
      ->setPath('./images/pokephp_logo.png') //definindo o  logo no topo do slide
      ->setHeight(36) //definindo a altura da forma
      ->setOffsetX(10) //definindo as coordenadas do eixo X referente a posição da forma
      ->setOffsetY(10); //definindo as coordenadas do eixo Y referente a posição da forma
$shape->getShadow()->setVisible(true) //definindo uma sombra na imagem
                   ->setDirection(45) //definindo a direção da sombra
                   ->setDistance(10); //definindo a distancia da sombra

$shape = $currentSlide->createRichTextShape() //criando uma forma (texto)
      ->setHeight(300) //definindo a altura da forma
      ->setWidth(600) //definindo a largura da forma
      ->setOffsetX(170) //definindo as coordenadas do eixo X referente a posição da forma
      ->setOffsetY(200); //definindo as coordenadas do eixo Y referente a posição da forma
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER ); //definindo o alinhamento do paragrafo
$textRun = $shape->createTextRun('Stand-Up do Pokemãobr!'); //definindo o texto a ser escrito
$textRun->getFont()->setBold(true) //definindo a fonte como negrito
                   ->setSize(60) //definindo o tamanho da fonte
                   ->setColor( new Color( 'FFE06B20' ) ); //definindo a cor da fonte

$shape = $currentSlide->createRichTextShape() //criando uma forma (texto)
      ->setHeight(100) //definindo a altura da forma
      ->setWidth(600) //definindo a largura da forma
      ->setOffsetX(10)//definindo as coordenadas do eixo X referente a posição da forma
      ->setOffsetY(640); //definindo as coordenadas do eixo Y referente a posição da forma
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT ); //definindo o alinhamento do paragrafo
$textRun = $shape->createTextRun('@pokemaobr'); //definindo o texto a ser escrito
$textRun->getFont()->setBold(true) //definindo a fonte como negrito
                   ->setSize(20) //definindo o tamanho da fonte
                   ->setColor( new Color( '555555' ) ); //definindo a cor da fonte

$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007'); //definindo o tipo de arquivo como PowerPoint2007
$oWriterPPTX->save(__DIR__ . "/exemplo1.pptx"); //definindo o nome de arquivo de saída
$oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation'); //definindo o tipo de arquivo como Open Document
$oWriterODP->save(__DIR__ . "/exemplo1.odp"); //definindo o nome de arquivo de saída