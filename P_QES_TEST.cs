//p.s. в using много мусора, многое не требуется
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EleWise.ELMA.API;
using EleWise.ELMA.Model.Common;
using EleWise.ELMA.Model.Entities;
using EleWise.ELMA.Model.Managers;
using EleWise.ELMA.Model.Types.Settings;
using EleWise.ELMA.Model.Entities.ProcessContext;
using Context = EleWise.ELMA.Model.Entities.ProcessContext.P_QES_TEST;
using EleWise.ELMA.Services;
using EleWise.ELMA.Scheduling;
using EleWise.ELMA.Documents.Models;
using EleWise.ELMA.ConfigurationModel;
using EleWise.ELMA.Model.Metadata;
using EleWise.ELMA.Documents.Models.Nomenclature;
using EleWise.ELMA.Model.Entities.EntityReferences;
using EleWise.ELMA.Model.Services;
using EleWise.ELMA.Tasks;
using EleWise.ELMA.Extensions;
using EleWise.ELMA.Files;
using Aspose.Words;
using EleWise.ELMA.Documents;
using EleWise.ELMA.Documents.Managers;
using EleWise.ELMA.Documents.Metadata;
using EleWise.ELMA.Security;
using EleWise.ELMA.Templates;
using EleWise.ELMA.Runtime.Managers;
using EleWise.ELMA.Common;
using EleWise.ELMA.Common.Models;
using Aspose.Cells;
using Aspose.Words.Saving;
using Aspose.Pdf;
using System.IO;
using System.Security.Cryptography;
using EleWise.ELMA.Security.Managers;
using EleWise.ELMA.Documents.Docflow;
using System.IO.Compression;
using ICSharpCode.SharpZipLib;
using NHibernate;
using System.Text.RegularExpressions;

namespace EleWise.ELMA.Model.Scripts
{
    partial class P_QES_TEST_Scripts //: EleWise.ELMA.Workflow.Scripts.ProcessScriptBase<Context>
	{

        /// <summary>
        /// Создание визуализированной версии
        /// </summary>
        /// <param name="context">Контекст процесса</param>
        public virtual void CreateVisualisedVersion(Context context)
        {
            //context.TemplateVersion - Версия шаблон
            //context.VisualisedVersion - Версия с визуализацией
            //context.SigFile - Файл с отделенной ЭП
            //context.Document - Документ ELMA


            //Инициализируем контейнер изображения
			var imgContainer = new QesImageContainer (null);
			//Копируем сразу файл с темплейта
			context.VisualisedVersion = PublicAPI.Services.File.CopyFile (context.TemplateVersion);
			var docAssignUser = context.Document.CurrentVersion.SignedUsers.LastOrDefault ();
			if(docAssignUser == null){
				docAssignUser = context.Document.Versions.Where(c=>c.SignedUsers.Any()).LastOrDefault().SignedUsers.LastOrDefault();
			}
			//Получаем подпись ЭП			
			var digiSign = EleWise.ELMA.Documents.Managers.DocumentVersionExtManager.GetDigitalSignature (docAssignUser, PublicAPI.Enums.Documents.DigitalSignature.SignatureGeneratingType.Content).Signature;
		
			System.Security.Cryptography.X509Certificates.X509Certificate2 sign = new System.Security.Cryptography.X509Certificates.X509Certificate2 (digiSign);
			
			//Запись SIG файла в контекст
			System.IO.File.WriteAllBytes(QesFileManager.folder + QesFileManager.sigFileName, digiSign);
			var sigFile = new EleWise.ELMA.Files.BinaryFile ();
			string name = Regex.Replace(context.Document.Name + ".sig", @"\s+", " ");
			sigFile.Name = string.Join("_", name.Split(Path.GetInvalidFileNameChars()));
			sigFile.InitializeContentFilePath();
			System.IO.File.Copy (QesFileManager.folder + QesFileManager.sigFileName, sigFile.ContentFilePath);
			PublicAPI.Services.File.SaveFile (sigFile);
			context.SigFile = sigFile;
			
			string Signer = "";
			Match match = Regex.Match(sign.Subject, @"SN=([^,]+),\sG=([^,]+)");
			if (match.Success)
			{
				Signer = match.Groups[1].Value + " " + match.Groups[2].Value;

			} else {
				throw new Exception("Произошла ошибка генерации штампа. RegEx не совпал с сертификатом по SN= ,G=");
			}


			//Формируем класс-контейнер с атрибутами ЭП
			var qesContainer = new QesContainer () {
				SerialNumber = sign.SerialNumber.ToUpper (),
				NotBefore = sign.NotBefore.ToShortDateString (),
				NotAfter = sign.NotAfter.ToShortDateString (),
				HolderFullName = Signer
			};
			//Создание файла штампа
			var stampFile = new EleWise.ELMA.Files.BinaryFile ();
			stampFile.Name = "QES_StampTemplate.docx";
			stampFile.InitializeContentFilePath ();
			//Copy шаблона штампа ЭП в binaryFile
			System.IO.File.Copy (QesFileManager.folder + QesFileManager.stampTemplateFileName, stampFile.ContentFilePath);
			PublicAPI.Services.File.SaveFile (stampFile);
			//Генерация штампа по атрибутам из контейнера
			PublicAPI.Services.DocumentGenerator.Generate (stampFile, qesContainer);
			PublicAPI.Services.File.SaveFile (stampFile);
			//Рендер в .png
			Aspose.Words.Document pngDoc = new Aspose.Words.Document (stampFile.ContentFilePath);
			Aspose.Words.Saving.ImageSaveOptions imageSaveOptions = new Aspose.Words.Saving.ImageSaveOptions (Aspose.Words.SaveFormat.Png);
            //Разрешение в DPI (в зависимости от размера еще скейлить будет)
			imageSaveOptions.Resolution = 150;
		
			//Страницы для рендера
			imageSaveOptions.PageSet = new Aspose.Words.Saving.PageSet (new Aspose.Words.Saving.PageRange (0, 0));
            //Скейлинг изображения. Например если слишком маленький шаблон штампа или наоборот
			imageSaveOptions.Scale = 1.5f;
			pngDoc.Save (QesFileManager.folder + QesFileManager.stampRenderedFileName, imageSaveOptions);
			//Формирование контейнера для вставки штампа в документ по шаблону
			//Для вставки в документ используем - Image
			
			imgContainer.Stamp.Name = QesFileManager.stampRenderedFileName;
			imgContainer.Stamp.InitializeContentFilePath ();
			System.IO.File.Copy (QesFileManager.folder + QesFileManager.stampRenderedFileName, imgContainer.Stamp.ContentFilePath);
			PublicAPI.Services.File.SaveFile (imgContainer.Stamp);
			//Генерируем в VisualisedVersion по датасурсу IMG контейнера
			PublicAPI.Services.DocumentGenerator.Generate (context.VisualisedVersion, imgContainer);
			PublicAPI.Services.File.SaveFile (context.VisualisedVersion);
			
			
			//Короче делаем PDF из него PDFA
			string curFileExt = context.VisualisedVersion.Extension.ToLower();
			//Удаляем файл PDF
			System.IO.File.Delete (QesFileManager.folder + QesFileManager.pdfFileName);
			//Удаляем файл PDFA
			System.IO.File.Delete (QesFileManager.folder + QesFileManager.pdfaFileName);
			//Удаляем файл который будет конвертироватся с учетом расширения
			System.IO.File.Delete (QesFileManager.folder + QesFileManager.toPdfFileName + curFileExt);
			//Копируем текущую версию в TEMP расположение
			System.IO.File.Copy (context.VisualisedVersion.ContentFilePath, QesFileManager.folder + QesFileManager.toPdfFileName + curFileExt);
			//Если файл .xlsx используем библиотеку ASPOSE.CELLS
			if (curFileExt == ".xlsx" || curFileExt == ".xls") {
				Aspose.Cells.Workbook doc = new Aspose.Cells.Workbook (QesFileManager.folder + QesFileManager.toPdfFileName + curFileExt);
				foreach (var worksheet in doc.Worksheets) {
					worksheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
					worksheet.PageSetup.Orientation = PageOrientationType.Landscape;
				}
				Aspose.Cells.PdfSaveOptions saveOptions = new Aspose.Cells.PdfSaveOptions (Aspose.Cells.SaveFormat.Pdf);
				saveOptions.AllColumnsInOnePagePerSheet = true;
				doc.Save (QesFileManager.folder + QesFileManager.pdfFileName, saveOptions);
				//Если файл .DOC или .DOCX используем библиотеку ASPOSE.WORDS
			}
			else if (curFileExt == ".docx" || curFileExt == ".doc") {
				Aspose.Words.Document doc = new Aspose.Words.Document (QesFileManager.folder + QesFileManager.toPdfFileName + curFileExt);
				doc.Save (QesFileManager.folder + QesFileManager.pdfFileName, Aspose.Words.SaveFormat.Pdf);
			} else
			{
                //вызываем исключение если у нас не docx/doc/xls/xlsx
                //Важно! С .odt и прочими опен офисовскими форматами работать не будет!
                //В теории должно работать со всеми майкрософт офис ПО'хами
				throw new Exception();
			}
			//Конвертируем PDF в PDF/A
			//Ремарка - PDF и PDF имеют одинаковое расширение файла, НО используется разные методы работы pdf
			//В примере используется PDF/A-1B, как я понимаю, это голый PDF без встроенного контейнера
			//upd. потестировал. Судя по размеру A-1B всё таки с контейнером. 
			//Хотя я без понятия. но типо pdf обычный меньше чем A-1B, но это логично, т.к. в PDF/A должны быть метаданные
			//для независимой от внешних факторов работы
			Aspose.Pdf.Document pdfADoc = new Aspose.Pdf.Document (QesFileManager.folder + QesFileManager.pdfFileName);
			pdfADoc.Convert (new System.IO.MemoryStream (), Aspose.Pdf.PdfFormat.PDF_A_1B, Aspose.Pdf.ConvertErrorAction.Delete);
			pdfADoc.Save (QesFileManager.folder + QesFileManager.pdfaFileName);
			//Создаем версию документа на основании сгенерированной PDF/A
			var convertedVersion = new EleWise.ELMA.Files.BinaryFile ();
			//Костыль что бы файл был pdf
			string nameConv = Regex.Replace(context.Document.Name + ".pdf", @"\s+", " ");
			convertedVersion.Name = string.Join("_", nameConv.Split(Path.GetInvalidFileNameChars()));
			//Создаем путь к binaryFile
			convertedVersion.InitializeContentFilePath ();
			//Копируем из шаблона в binaryFile
			System.IO.File.Copy (QesFileManager.folder + QesFileManager.pdfaFileName, convertedVersion.ContentFilePath);
			//Сохраняем файл в БД
			PublicAPI.Services.File.SaveFile (convertedVersion);
			context.VisualisedVersion = convertedVersion;
            //Устанавливаем текущей - бинго
			PublicAPI.Docflow.DocumentVersion.AddDocumentVersion (context.Document, context.VisualisedVersion, PublicAPI.Enums.Documents.DocumentVersionStatus.Current);
        }
        
        /// <summary>
        /// Создание зип-архива
        /// </summary>
        /// <param name="context">Контекст процесса</param>
        public virtual void CreateZip(Context context)
        {
			//Удаляем старый ZIP архив
			//p.s. я устал создавать константы в классе под название файлов
			System.IO.File.Delete (QesFileManager.folder + "sign.zip");
			
			//Файлы для помещения в архив
			List<EleWise.ELMA.Files.BinaryFile> filesToCompress = new List<EleWise.ELMA.Files.BinaryFile>();
			
			string nameSource = Regex.Replace(context.Document.Name, @"\s+", " ");
			string name = string.Join("_", nameConv.Split(Path.GetInvalidFileNameChars()));
			context.PDFForSignance.Name = string.Join("_", (name+".pdf").Split(Path.GetInvalidFileNameChars()));
			context.SigFile.Name = string.Join("_", (name+".sig").Split(Path.GetInvalidFileNameChars()));
			filesToCompress.Add(context.PDFForSignance);
			filesToCompress.Add(context.SigFile);
			//использование одной единственной либы для зип архивов в элме.
			//p.s. Спасибо EleWise, что вкорячили в элму Aspose.Notes и прочее ОчЕнЬ ПоЛеЗнОе, но только не Aspose.ZIP .NET 
			using (ICSharpCode.SharpZipLib.Zip.ZipOutputStream archiveStream = new ICSharpCode.SharpZipLib.Zip.ZipOutputStream(System.IO.File.Create(QesFileManager.folder + "sign.zip"))){
				//Устанавливаем уровень сжатия 0 (максимальный - 9), для предотвращения искажения в ходе компрессии
				archiveStream.SetLevel(0); 
				
				//Хз че это, тут какие то сложные действия с потоками, скопировано гордо со стак оверфлов
				byte[] buffer = new byte[4096]; 
				foreach(var file in filesToCompress)
				{
					//Entry (файл для помещения в архив)
					var entry = new ICSharpCode.SharpZipLib.Zip.ZipEntry(file.Name);
					entry.IsUnicodeText = true;
					
					//Дата последнего изменения архива
					entry.DateTime = DateTime.Now;
					//Говорит потоку, о том что ща будет новый файл
        			archiveStream.PutNextEntry(entry);
        			//читаем файл
        			using (FileStream fs = System.IO.File.OpenRead(file.ContentFilePath)) 
        			{
        				//Какой то мэджик с чтением побайтово. Зачем так? я не знаю...
        				int sourceBytes;
			            do
			            {
			            sourceBytes = fs.Read(buffer, 0, buffer.Length);
			            archiveStream.Write(buffer, 0, sourceBytes);
			            } while (sourceBytes > 0);
			        }
				}
				//Завершаем работу потока и закрываем его
				archiveStream.Finish();
				archiveStream.Close();
			}
			//Формирование байнари файла в эльме
			var zipBinaryFile = new EleWise.ELMA.Files.BinaryFile();
			zipBinaryFile.Name = "sign.zip";
			zipBinaryFile.InitializeContentFilePath();
			System.IO.File.Copy (QesFileManager.folder + "sign.zip", zipBinaryFile.ContentFilePath);
			PublicAPI.Services.File.SaveFile (zipBinaryFile);
			context.SignedArchive = zipBinaryFile;
        }

		/// <summary>
		/// Класс менеджер настроек файлов для ЭЦП
		/// </summary>
		public static class QesFileManager
		{
			/// <summary>
			/// Папка хранящая ЭЦП
			/// </summary>
			public const string folder = @"c:\\ELMA_QES\";

			/// <summary>
			/// Наименование файла-холдера версии pdfa
			/// </summary>
			public const string pdfaFileName = "PDFAVersion.pdf";
			
			/// <summary>
			/// Наименование файла-холдера версии pdfa с визуализацией 
			/// </summary>
			public const string pdfaVisFileName = "PDFAVisualisedVersion.pdf";

			/// <summary>
			/// Наименование файла-холдера версии pdf
			/// </summary>
			public const string pdfFileName = "PDFVersion.pdf";

			/// <summary>
			/// Наименование файла холдера для конвертации в pdf без расширения
			/// </summary>
			public const string toPdfFileName = "ToPDFVersion";

			/// <summary>
			/// Наименование файла холдер штампа
			/// </summary>
			public const string stampTemplateFileName = "StampTemplate.docx";

			/// <summary>
			/// Наименование файла рендера штампа
			/// </summary>
			public const string stampRenderedFileName = "StampRendered.png";

			public const string stampDemoFileName = "StampDemo.png";
			
			/// <summary>
			/// Наименование Sig файла
			/// </summary>
			public const string sigFileName = "ЭП.sig";
		}


		/// <summary>
		/// Класс контейнер штампа для генерации по шаблону в документ
		/// </summary>
		public class QesImageContainer
		{
			public EleWise.ELMA.Files.BinaryFile Stamp {
				get;
				set;
			}
			
			public string Date
			{
				get;
				set;
			}
			public string RegNum
			{
				get;
				set;
			}
			

			public QesImageContainer (EleWise.ELMA.Files.BinaryFile? StampFile)
			{
				if (StampFile != null) {
					this.Stamp = StampFile;
				}
				else {
					this.Stamp = new EleWise.ELMA.Files.BinaryFile ();
				}
			}
		}

		/// <summary>
		/// Класс контейнер атрибутов ЭП, для генерации по шаблону штампа
		/// </summary>
		public class QesContainer
		{
			/// <summary>
			/// Серийный номер ЭП
			/// </summary>
			public string SerialNumber {
				get;
				set;
			}

			/// <summary>
			/// Срок действия ЭП с даты
			/// </summary>
			public string NotBefore {
				get;
				set;
			}

			/// <summary>
			/// Срок действия ЭП до даты
			/// </summary>
			public string NotAfter {
				get;
				set;
			}

			/// <summary>
			/// Держатель ЭП (ФИО)
			/// </summary>
			public string HolderFullName {
				get;
				set;
			}
		}


    }
}