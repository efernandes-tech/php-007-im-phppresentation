<?php

// Autoload do composer.
require_once 'vendor/autoload.php';

// Classe do PhpPresentation.
use PhpOffice\PhpPresentation\IOFactory;
// Classe para manipular os arquivos.
use PhpOffice\PhpPresentation\PhpPresentation;
// Classe de estilo de cores.
use PhpOffice\PhpPresentation\Style\Alignment;
// Classe de estilo de alinhamentos.
use PhpOffice\PhpPresentation\Style\Color;

// Instanciando uma nova apresentação.
$objPHPPowerPoint = new PhpPresentation();

// Retornando o slide ativo.
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Criando uma forma de desenho (imagem).
$shape = $currentSlide->createDrawingShape();
// Definindo o nome da forma.
$shape->setName('PokePHP Logo')
// Definindo a descrição da forma.
    ->setDescription('PokePHP Logo')
// Definindo o  logo no topo do slide.
    ->setPath('./images/pokephp_logo.png')
// Definindo a altura da forma.
    ->setHeight(36)
// Definindo as coordenadas do eixo X referente a posição da forma.
    ->setOffsetX(10)
// Definindo as coordenadas do eixo Y referente a posição da forma.
    ->setOffsetY(10);
// Definindo uma sombra na imagem.
$shape->getShadow()->setVisible(true)
// Definindo a direção da sombra.
    ->setDirection(45)
// Definindo a distancia da sombra.
    ->setDistance(10);

// Criando uma forma (texto).
$shape = $currentSlide->createRichTextShape()
// Definindo a altura da forma.
    ->setHeight(300)
// Definindo a largura da forma.
    ->setWidth(600)
// Definindo as coordenadas do eixo X referente a posição da forma.
    ->setOffsetX(170)
// Definindo as coordenadas do eixo Y referente a posição da forma.
    ->setOffsetY(200);
// Definindo o alinhamento do paragrafo.
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
// Definindo o texto a ser escrito.
$textRun = $shape->createTextRun('Stand-Up do Pokemãobr!');
// Definindo a fonte como negrito.
$textRun->getFont()->setBold(true)
// Definindo o tamanho da fonte.
    ->setSize(60)
// Definindo a cor da fonte.
    ->setColor(new Color('FFE06B20'));

// Criando uma forma (texto).
$shape = $currentSlide->createRichTextShape()
// Definindo a altura da forma.
    ->setHeight(100)
// Definindo a largura da forma.
    ->setWidth(600)
// Definindo as coordenadas do eixo X referente a posição da forma.
    ->setOffsetX(10)
// Definindo as coordenadas do eixo Y referente a posição da forma.
    ->setOffsetY(640);
// Definindo o alinhamento do paragrafo.
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
// Definindo o texto a ser escrito.
$textRun = $shape->createTextRun('@pokemaobr');
// Definindo a fonte como negrito.
$textRun->getFont()->setBold(true)
// Definindo o tamanho da fonte.
    ->setSize(20)
// Definindo a cor da fonte.
    ->setColor(new Color('555555'));

// Definindo o tipo de arquivo como PowerPoint2007.
$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
// Definindo o nome de arquivo de saída.
$oWriterPPTX->save(__DIR__ . "/exemplo1.pptx");
// Definindo o tipo de arquivo como Open Document.
$oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation');
// Definindo o nome de arquivo de saída.
$oWriterODP->save(__DIR__ . "/exemplo1.odp");
