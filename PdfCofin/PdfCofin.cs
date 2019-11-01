using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace PdfCofin
{
    class Eventos : PdfPageEventHelper
    {
        public Font fonte { get; set; }

        public Eventos(Font fonte_)
        {
            fonte = fonte_;
        }


        // Este método cria um cabeçalho para o documento
        public override void OnStartPage(PdfWriter writer, Document document)
        {
            // Cria um novo paragrafo com o texto do cabeçalho
            //Paragraph ph = new Paragraph();

            //Criar 1º Tabela
            PdfPTable table = new PdfPTable(2); //Adicionar 2 Celulas
            table.HorizontalAlignment = 0;
            table.TotalWidth = 500f;
            table.LockedWidth = true;
            float[] widths = new float[] { 20f, 80f }; //Tamanho das celulas na tabela
            table.SetWidths(widths);
            table.DefaultCell.Border = Rectangle.NO_BORDER;

            //1º Celula - da 1º Tabela
            Image image = Image.GetInstance(@"C:\Users\xxx\Downloads\Brasao.jpg");
            PdfPCell cell = new PdfPCell(image);
            image.Alignment = Element.ALIGN_CENTER;
            image.ScaleAbsolute(80, 80);
            cell.Border = 0;
            cell.AddElement(image); //Adicionar imagem na celula
            table.AddCell(cell); //Adicionar celula na tabela

            //2º Tabela cabeçalho - está dentro da 1º tabela
            PdfPTable tabl = new PdfPTable(1);
            tabl.DefaultCell.Border = Rectangle.NO_BORDER;

            //Textos 2º Tabela
            Phrase p1 = new Phrase("SECRETARIA  DE ESTADO DOS NEGÓCIOS DA SEGURANÇA PÚBLICA", FontFactory.GetFont("Arial", 9, Font.BOLD));
            Phrase p2 = new Phrase("POLICIA MILITAR DO ESTADO DE SÃO PAULO", FontFactory.GetFont("Arial", 9, Font.BOLD));
            Phrase p3 = new Phrase("Diretoria de Finanças e Patrimônio", FontFactory.GetFont("Arial", 8, Font.NORMAL));
            Phrase p4 = new Phrase("COFIN - Controle orçamentário e Financeiro", FontFactory.GetFont("Arial", 8, Font.NORMAL));
            Phrase p5 = new Phrase("DIÁRIA DE DILIGÊNCIA", FontFactory.GetFont("Arial", 9, Font.BOLD));
            //Phrase p6 = new Phrase("2019DE15900032", FontFactory.GetFont("Arial", 9, Font.BOLD));

            //Adicionar celulas na 2º tabela
            tabl.AddCell(new PdfPCell(p1) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
            tabl.AddCell(new PdfPCell(p2) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
            tabl.AddCell(new PdfPCell(p3) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
            tabl.AddCell(new PdfPCell(p4) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
            tabl.AddCell(new PdfPCell(p5) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
            //tabl.AddCell(ce6);
            table.AddCell(tabl);

            //Adicionar tabela no documento
            document.Add(table);
        }

        //public override void OnEndPage(PdfWriter writer, Document document)
        public override void OnEndPage(PdfWriter writer, Document _pdfDoc)
        {
            Document document = new Document();

            PdfContentByte cb;
            PdfTemplate template;
            BaseFont bf = null;
            BaseFont b = null;
            //DateTime PrintTime = DateTime.Now;
            //PrintTime = DateTime.Now;
            bf = BaseFont.CreateFont(BaseFont.HELVETICA_OBLIQUE, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            b = BaseFont.CreateFont(BaseFont.HELVETICA_BOLDOBLIQUE, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb = writer.DirectContent;
            template = cb.CreateTemplate(50, 50);

            base.OnEndPage(writer, document);
            //int pageN = writer.PageNumber;
            String text = "'Nós, Policiais Militares, sob a proteção de Deus, estamos compromissados com a defesa da Vida, da Integridade Física e da Dignidade da Pessoa Humana.'";
            float len = bf.GetWidthPoint(text, 6);
            Rectangle pageSize = document.PageSize;
            cb.SetRGBColorFill(0, 0, 0);
            cb.BeginText();
            cb.SetFontAndSize(b, 6);
            cb.SetTextMatrix(pageSize.GetLeft(75), pageSize.GetBottom(15));
            cb.ShowText(text);
            cb.EndText();
            cb.AddTemplate(template, pageSize.GetLeft(295) + len, pageSize.GetBottom(30));
            cb.BeginText();
            cb.SetFontAndSize(bf, 7);
            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER,
            "Documento assinado eletronicamente, informações registradas nos servidores da PMESP.",
            pageSize.GetRight(295),
            pageSize.GetBottom(35), 0);
            cb.EndText();
        }
    }
    class PdfCofin
    {
        Document _pdfDoc; //Objeto principal interno
        PdfWriter _pdfWriter; //Objeto principal interno
        private Diaria _diaria;
        private Font font8N;
        private Font font8;
        private Font font9;
        Paragraph getParagraph(IList<Phrase> ph)
        {
            Paragraph prCofin = new Paragraph();
            prCofin.AddRange(ph);
            return prCofin;
        }

        public PdfCofin() //Contrutor com parametros
        {
            font8N = FontFactory.GetFont("Arial", 8, Font.NORMAL);
            font8 = FontFactory.GetFont("Arial", 8, Font.BOLD);
            font9 = FontFactory.GetFont("Arial", 9, Font.BOLD);

            Document doc = new Document(PageSize.A4);
            doc.SetMargins(40, 40, 40, 80);
            doc.AddCreationDate();
            string caminho = @"C:\Users\xxx\Downloads\" + "File.pdf";
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(caminho, FileMode.Create));

            _pdfDoc = doc; //Instanciar para quem chamar
            _pdfWriter = writer; //Instanciar para quem chamar // cria um objeto do tipo FontFamily, que contem as propriedades de uma fonte

            Font.FontFamily familha = new Font.FontFamily();// atribui a familia da fonte, no caso Courier
            familha = iTextSharp.text.Font.FontFamily.COURIER;  // cria uma fonte atribuindo a familha, o tamanho da fonte e o estilo (normal, negrito...)
            Font fonte = new Font(familha, 8); // cria uma instancia da classe eventos, é uma classe que mostrarei posteriormente

            Eventos ev = new Eventos(fonte); // esta clase trata a criação do cabeçalho e rodapé da página
            _pdfWriter.PageEvent = ev;  // seta o atributo de eventos da classe com a variavel de eventos criada antes

            fonte = new Font(familha, 8); // altera a fonte para normal, a negrito era apenas para o cabeçalho e rodapé da página

            _pdfDoc.Open();
        }
        public void Page() //Metodo para criar o pagina
        {
            //3º tabela
            PdfPTable tabSolicitante = new PdfPTable(3);
            tabSolicitante.HorizontalAlignment = 0;
            tabSolicitante.TotalWidth = 500f;
            tabSolicitante.LockedWidth = true;
            float[] widths = new float[] { 80f, 23f, 35f }; //Tamanho das celulas na tabela
            tabSolicitante.SetWidths(widths);
            tabSolicitante.DefaultCell.Border = Rectangle.NO_BORDER;

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font9)}))
            {
                Colspan = 3,
                HorizontalAlignment = Element.ALIGN_RIGHT,
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font9) }))
            {
                Colspan = 3,
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Solicitante: ", font8N),
              new Phrase("XXX", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CPF: ", font8N),
              new Phrase("xxx.xxx.xxx-xx", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("RG: ", font8N),
              new Phrase("xx.xxx.xxx-x", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Função: ", font8N),
              new Phrase("ANALISTA DE SISTEMAS", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font9) }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("COD. OPM: ", font8N),
              new Phrase("xxx", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Endereço: ", font8N),
              new Phrase("RUA DOS BOBOS, Nº 0", font8 )
            }))
            {
                Colspan = 3,
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Banco: ", font8N),
              new Phrase("xxx", font8 ),
              new Phrase("                        Ag: ", font8N),
              new Phrase("xxx", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("C/C: ", font8N),
              new Phrase("xxx", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Retribuição Mensal: ", font8N),
              new Phrase("R$ 0,00", font8 )
            }))
            {
                Border = 0
            });

            tabSolicitante.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Informação Contábil: ", font8N),
              new Phrase(" ", font8 )
            }))
            {
                Colspan = 3,
                Border = 0
            });

            _pdfDoc.Add(tabSolicitante);

            //4º tabela
            PdfPTable tab4 = new PdfPTable(10);
            tab4.HorizontalAlignment = 0;
            tab4.SpacingBefore = 7f;
            tab4.TotalWidth = 500f;
            tab4.LockedWidth = true;
            float[] width = new float[] { 50f, 15f, 10f, 8f, 10f, 10f, 7f, 7f, 10f, 15f }; //Tamanho das celulas na tabela
            tab4.SetWidths(width);

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Motivo do Deslocamento: ", font8),
              new Phrase("nova diária para teste do sistema de pagamento", font8N )
            }))
            {
                Colspan = 10,
                Border = 0,
                BorderWidthBottom = 0
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Municípo de destino", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Data", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Refeiç", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Pous.", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("% art4º", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("% art5º", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Emp.", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Pern.", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("% local", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Valor", font8) }))
            {
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("SAO PAULO-SP (Km: 330 - Pop 12.106.920)", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235)
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("01/09/2019", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("100%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("X", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("X", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("80%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("0,00", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("SAO PAULO-SP (Km: 330 - Pop 12.106.920)", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235)
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/09/2019", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("100%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("X", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("X", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("80%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("0,00", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("SAO PAULO-SP (Km: 330 - Pop 12.106.920)", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235)
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("03/09/2019", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("100%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("X", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("X", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("80%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("0,00", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("DIA DE RETORNO", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235)
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("04/09/2019", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("40%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("80%", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("0,00", font8N) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Observação: ", font8N),
              new Phrase("Partida da origem às 7h00, saída do destino às 19h00 e retorno à origem ás 23h00", font8N )
            }))
            {
                Colspan = 8,
                BorderWidthRight = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("TOTAL: ", font8) }))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>()
            { new Phrase("R$0,00", font8N)}))
            {
                BorderWidthLeft = 0,
                BorderColorBottom = new BaseColor(235, 235, 235),
                BorderColorLeft = new BaseColor(235, 235, 235),
                BorderColorTop = new BaseColor(235, 235, 235),
                BorderColorRight = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("VALOR TOTAL A RECEBER: ", font8) }))
            {
                Colspan = 9,
                Border = 0,
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            tab4.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("R$0,00", font8) }))
            {
                Border = 0,
                HorizontalAlignment = Element.ALIGN_RIGHT
            });

            _pdfDoc.Add(tab4);

            //5º tabela
            PdfPTable tab5 = new PdfPTable(2);
            tab5.HorizontalAlignment = 0;
            tab5.SpacingBefore = 7f;
            tab5.TotalWidth = 500f;
            tab5.LockedWidth = true;
            float[] tam = new float[] { 100f, 55f }; //Tamanho das celulas na tabela
            tab5.SetWidths(tam);

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Memorial de cálculo", font8) }))
            {
                Colspan = 2,
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Base Legal: Dec. 48292 de 02DEZ03 e Dec. 61.397 de 04AGO15", font8) }))
            {
                Colspan = 2,
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Valor UFESP DE 01 de janeiro a 31 de dezembro de 2019", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("R$ 0,00", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Quantidade de UFESP: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("7", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Valor da diária (Quantidade de UFESP x Valor da UFESP)", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("R$ 0,00", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Quantidade de pernoite: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("0", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Quantidade de dia empenhado: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("0", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Fórmula valor: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("(Valor da diária + % art.3º)x %art. 5º", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Fórmula valor quando equipe de apoio Governador: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("(Valor da diária + % art.3º + % art.4º)x %art. 5º", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Dia de retorno (regresso à sede): ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("SIM", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Fórmula do retorno quando a chegada à sede ocorrer a partir das 13h e antes das 19h: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("(Valor da diária + %art.3º) x 20%", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Fórmula do retorno quando a chegada à sede ocorrer a partir das 19h: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("(Valor da diária + %art.3º) x 40%", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Valor de diárias recebidas no mês: ", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("R$ 0,00", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Fórmula da Glosa: ", font8N),
              new Phrase("Valor a receber + Valor recebido no mês - (Retribuição mensal/2)", font8 )
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                Colspan = 2
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("(Fonte de Quilometragem : Interface entre aplicativo e programação - API Google Maps)", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                Colspan = 2
            });

            tab5.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("(Fonte de população: Tabela do IBGE atualizada em 2010)", font8) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                Colspan = 2
            });

            _pdfDoc.Add(tab5);

            //6º tabela
            PdfPTable tab6 = new PdfPTable(3);
            tab6.HorizontalAlignment = 0;
            tab6.SpacingBefore = 17f;
            tab6.TotalWidth = 500f;
            tab6.LockedWidth = true;
            float[] tama = new float[] { 15f, 100f, 25f }; //Tamanho das celulas na tabela
            tab6.SetWidths(tama);

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Data", font8) }))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Responsável/Histórico", font8) }))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Status", font8) }))
            {
                HorizontalAlignment = Element.ALIGN_CENTER,
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 10:49:08", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CEL PM: ", font8),
              new Phrase("Diária Aguardando Recursos - IdCofin: xxx (xxx)", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Recursos", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 10:49:08", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CEL PM: ", font8),
              new Phrase("Autorizo o pagamento da Diária de Diligência em concordância com o art. 15 do Dec.  48.292 de 02DEZ03", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Recursos", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 10:17:19", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CAP PM: ", font8),
              new Phrase("Solicitação em conformidade com a legislação vigente e apta ao pagamento.", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Análise do Dirigente", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 09:56:04", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("MAJ PM: ", font8),
              new Phrase("Atesto que a Escala de Serviço e o Controle de Frequência estão em conformidade e que não foram sacadas Diárias de Alimentação nos dias relativos à solicitação de Diárias de Diligência, conforme previsto no art. 1º do Dec. 59.609, de 16OUT13. Encaminhamento à UGE responsável pelo pagamento.", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Análise da UGE", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 09:55:29", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("MAJ PM: ", font8),
              new Phrase("Aprovo esta solicitação. Atesto o cumprimento da Escala de Serviço e confirmo solidariamente os direitos solicitados pelo 2.SGT PM, conforme preceitua o art. 14 do Dec. 48.292 de 02DEZ03. Encaminho a solicitação ao MAJ PM, responsável pelo lançamento das Diárias de Alimentação no SIPA.", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Análise do Oficial Responsável SIPA", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 09:51:43", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("2. SGT PM: ", font8),
              new Phrase("Solicitação encaminhada a Sra.: MAJOR PM", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Análise do Superior Imediato", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 09:51:42", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("2. SGT PM: ", font8),
              new Phrase("Aceito o Termo de Ciência e Responsabilidade e declaro que as informações são expressão da verdade e estão de acordo com o Decreto nº 48.292 de 02DEZ03.", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Análise do Superior Imediato", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("02/SET/2019 09:51:42", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("2. SGT PM: ", font8),
              new Phrase("Aceito o Termo de Ciência e Responsabilidade e declaro que as informações são expressão da verdade e estão de acordo com o Decreto nº 48.292 de 02DEZ03.", font8N)
            }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            tab6.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Aguardando Análise do Superior Imediato", font8N) }))
            {
                BorderColor = new BaseColor(235, 235, 235),
                HorizontalAlignment = Element.ALIGN_CENTER,
                BackgroundColor = new BaseColor(235, 235, 235)
            });

            _pdfDoc.Add(tab6);

            //7º tabela
            PdfPTable tab7 = new PdfPTable(3);
            tab7.HorizontalAlignment = 0;
            tab7.SpacingBefore = 8f;
            tab7.TotalWidth = 500f;
            tab7.LockedWidth = true;
            float[] tamanho = new float[] { 50f, 50f, 50f }; //Tamanho das celulas na tabela
            tab7.SetWidths(tamanho);

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("ORDEM BANCÁRIA - OB", font9) }))
            {
                Border = 0,
                Colspan = 3,
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8) }))
            {
                Colspan = 3,
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("Nº DOC.: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("DATA EMISSÃO: ", font8),
              new Phrase("10SET2019 10:03", font8N)
            }))
            {
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("DATA LANCAMENTO: ", font8),
              new Phrase("10SET2019 10:03", font8N)
            }))
            {
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("UG: ", font8),
              new Phrase("xxx-SECRETARIA DA SEGURANCA PUBLICA", font8N)
            }))
            {
                Colspan = 2,
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("GESTÃO: ", font8),
              new Phrase("xxx  -ADMINIST.DIRETA", font8N)
            }))
            {
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("PD/NL/OC/LISTA", font8),
              new Phrase("xxx zzz", font8N)
            }))
            {
                Colspan = 2,
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("TIPO OB: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab7.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("STATUS: ", font8),
              new Phrase("*PAGAMENTO NORMAL*", font8N)
            }))
            {
                Colspan = 3,
                Border = 0,
            });

            _pdfDoc.Add(tab7);

            //8º tabela
            PdfPTable tab8 = new PdfPTable(3);
            tab8.HorizontalAlignment = 0;
            tab8.SpacingBefore = 8f;
            tab8.TotalWidth = 500f;
            tab8.LockedWidth = true;
            float[] ta = new float[] { 50f, 50f, 50f }; //Tamanho das celulas na tabela
            tab8.SetWidths(ta);

            tab8.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("DOM.BANC.EMITENTE", font9) }))
            {
                Border = 0,
                Colspan = 3,
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab8.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font9) }))
            {
                Border = 0,
                Colspan = 3,
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab8.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("BANCO: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab8.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("AGENCIA: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab8.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CONTA CORRENTE: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            _pdfDoc.Add(tab8);

            //9º tabela
            PdfPTable tab9 = new PdfPTable(3);
            tab9.HorizontalAlignment = 0;
            tab9.SpacingBefore = 8f;
            tab9.TotalWidth = 500f;
            tab9.LockedWidth = true;
            float[] t = new float[] { 50f, 50f, 50f }; //Tamanho das celulas na tabela
            tab9.SetWidths(t);

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("FAVORECIDO/DOMICILIO BANCARIO", font9) }))
            {
                Border = 0,
                Colspan = 3,
                HorizontalAlignment = Element.ALIGN_CENTER
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font9) }))
            {
                Border = 0,
                Colspan = 3,
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CNPJ/CPF/UG: ", font8),
              new Phrase("xxx - PROFESSOR XAVIER", font8N)
            }))
            {
                Border = 0,
                Colspan = 3
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("GESTÃO: ", font8),
              new Phrase(" ", font8N)
            }))
            {
                Border = 0,
                Colspan = 3
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("BANCO: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("AGENCIA: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CONTA CORRENTE: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("PROCESSO: ", font8),
              new Phrase("xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("VALOR: ", font8),
              new Phrase("0,00", font8N)
            }))
            {
                Border = 0,
            });

            tab9.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("FINALIDADE: ", font8),
              new Phrase("PAGAMENTO DE DIARIAS", font8N)
            }))
            {
                Border = 0,
            });

            _pdfDoc.Add(tab9);

            //10º tabela
            PdfPTable tab10 = new PdfPTable(6);
            tab10.HorizontalAlignment = 0;
            tab10.SpacingBefore = 8f;
            tab10.TotalWidth = 500f;
            tab10.LockedWidth = true;
            float[] amanho = new float[] { 50f, 50f, 50f, 50f, 50f, 50f }; //Tamanho das celulas na tabela
            tab10.SetWidths(amanho);

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("EVENTO", font9) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("REC/DESP", font9) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("INSC. DO EVENTO", font9) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("CLASSIF", font9) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("FONTE", font9) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("VALOR", font9) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("00,00", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase(" ", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("xxx", font8N) }))
            {
                Border = 0,
            });

            tab10.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("00,00", font8N) }))
            {
                Border = 0,
            });

            _pdfDoc.Add(tab10);

            //11º tabela
            PdfPTable tab11 = new PdfPTable(2);
            tab11.HorizontalAlignment = 0;
            tab11.SpacingBefore = 25f;
            tab11.TotalWidth = 500f;
            tab11.LockedWidth = true;
            float[] manho = new float[] { 50f, 50f }; //Tamanho das celulas na tabela
            tab11.SetWidths(manho);

            tab11.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("SITUACAO: ", font8),
              new Phrase("RELACIONADA - NUMERO: xxx", font8N)
            }))
            {
                Border = 0,
                Colspan = 2
            });

            tab11.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("LANCADO POR: ", font8),
              new Phrase("THOMAS WAYNE - xxx", font8N)
            }))
            {
                Border = 0,
            });

            tab11.AddCell(new PdfPCell(getParagraph(new List<Phrase>
            { new Phrase("DATA LANCAMENTO: ", font8),
              new Phrase("27/09/2019 15:40", font8N)
            }))
            {
                Border = 0,
            });

            _pdfDoc.Add(tab11);
            _pdfDoc.Close();
        }
    }
}