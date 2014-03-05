package yijon;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Locale;

import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NoteRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.RKRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

/**
 * 
 * 事件机制，读取Excel 97-2003 工作簿(*.xls)
 * 
 * POI开源项目实现
 * 
 * @author yijon 2011-11-18 下午02:46:06
 *
 */
public class ReadXlsEventModel implements HSSFListener {

	private int minColumns;
	private POIFSFileSystem fs;
	private PrintStream output;

	private int lastRowNumber;
	private int lastColumnNumber;

	/** 
	 * 遇到公式单元格的时候，是取值？还是获取公式字符串？ 
	 * true：取值，false：获取公式字符串
	 */
	private boolean outputFormulaValues = true;

	/** 为解析公式做准备 */
	private SheetRecordCollectingListener workbookBuildingListener;
	private HSSFWorkbook stubWorkbook;

	//处理记录的过程中，需要用到的数据对象
	private SSTRecord sstRecord;
	
	/**
	 * 定义，处理过程中，数据国际化的格式
	 */
	private FormatTrackingHSSFListener formatListener;
	
	/**
	 * 当前正在处理的sheet的索引
	 */
	private int sheetIndex = -1;
	private BoundSheetRecord[] orderedBSRs;
	private ArrayList boundSheetRecords = new ArrayList();

	// 获取公式字符串的准备
	private int nextRow;
	private int nextColumn;
	private boolean outputNextStringRecord;	
	
	/**
	 * Creates a new XLS -> CSV converter
	 * @param fs The POIFSFileSystem to process
	 * @param output The PrintStream to output the CSV to
	 * @param minColumns The minimum number of columns to output, or -1 for no minimum
	 */
	public ReadXlsEventModel(POIFSFileSystem fs, PrintStream output, int minColumns) {
		this.fs = fs;
		this.output = output;
		this.minColumns = minColumns;
	}

	/**
	 * Creates a new XLS -> CSV converter
	 * @param filename The file to process
	 * @param minColumns The minimum number of columns to output, or -1 for no minimum
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	public ReadXlsEventModel(String filename, int minColumns) throws IOException, FileNotFoundException {
		this(
				new POIFSFileSystem(new FileInputStream(filename)),
				System.out, minColumns
		);
	}
	
	/**
	 * 启动Excel文件的读写
	 */
	public void process() throws IOException {
		//定义“遇到空白单元格，自动跳过”，这样工作方式的监听器。
		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
		
		//处理过程中，需要添加国际化处理
		//构造函数如果不指定Locale，按照系统默认语言环境处理
		//如果需要显示特殊的日期或货币格式，可以设置需要的Locale
		//例如设置成简体中文格式：formatListener = new FormatTrackingHSSFListener(listener,Locale.CHINA);
		
		//定义为默认当前系统语言环境
		formatListener = new FormatTrackingHSSFListener(listener);

		//事件工厂
		HSSFEventFactory factory = new HSSFEventFactory();
		//工厂类必要参数
		HSSFRequest request = new HSSFRequest();

		//判断公式的处理方式
		if(outputFormulaValues) {
			request.addListenerForAllRecords(formatListener);
		} else {
			workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
			request.addListenerForAllRecords(workbookBuildingListener);
		}

		factory.processWorkbookEvents(request, fs);
	}
	
	/**
	 * Main HSSFListener method, processes events
	 * 每一个Excel单元格读取的时候，都会自动被触发该方法
	 */
	@Override
	public void processRecord(Record record) {
		int thisRow = -1;
		int thisColumn = -1;
		String thisStr = null;
		
		//依据不同的单元格类型，处理相应的业务逻辑
		//本例子当中的代码为官方示例，逻辑为打印输出
		//record.getSid()的所有类型，请参阅“Record 类型的 与 sid 的关系.txt”
		switch (record.getSid()){
			case BoundSheetRecord.sid:
				boundSheetRecords.add(record);
				break;
			case BOFRecord.sid:
				BOFRecord br = (BOFRecord)record;
				if(br.getType() == BOFRecord.TYPE_WORKSHEET) {
					// Create sub workbook if required
					if(workbookBuildingListener != null && stubWorkbook == null) {
						stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
					}
					
					// Output the worksheet name
					// Works by ordering the BSRs by the location of
					//  their BOFRecords, and then knowing that we
					//  process BOFRecords in byte offset order
					sheetIndex++;
					if(orderedBSRs == null) {
						orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
					}
					output.println();
					output.println( 
							orderedBSRs[sheetIndex].getSheetname() +
							" [" + (sheetIndex+1) + "]:"
					);
				}
				break;
	
			case SSTRecord.sid:
				sstRecord = (SSTRecord) record;
				break;
	
			case BlankRecord.sid:
				BlankRecord brec = (BlankRecord) record;
	
				thisRow = brec.getRow();
				thisColumn = brec.getColumn();
				thisStr = "";
				break;
			case BoolErrRecord.sid:
				BoolErrRecord berec = (BoolErrRecord) record;
	
				thisRow = berec.getRow();
				thisColumn = berec.getColumn();
				thisStr = "";
				break;
	
			case FormulaRecord.sid:
				FormulaRecord frec = (FormulaRecord) record;
	
				thisRow = frec.getRow();
				thisColumn = frec.getColumn();
	
				if(outputFormulaValues) {
					if(Double.isNaN( frec.getValue() )) {
						// Formula result is a string
						// This is stored in the next record
						outputNextStringRecord = true;
						nextRow = frec.getRow();
						nextColumn = frec.getColumn();
					} else {
						thisStr = formatListener.formatNumberDateCell(frec);
					}
				} else {
					thisStr = '"' +
						HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
				}
				break;
			case StringRecord.sid:
				if(outputNextStringRecord) {
					// String for formula
					StringRecord srec = (StringRecord)record;
					thisStr = srec.getString();
					thisRow = nextRow;
					thisColumn = nextColumn;
					outputNextStringRecord = false;
				}
				break;
	
			case LabelRecord.sid:
				LabelRecord lrec = (LabelRecord) record;
	
				thisRow = lrec.getRow();
				thisColumn = lrec.getColumn();
				thisStr = '"' + lrec.getValue() + '"';
				break;
			case LabelSSTRecord.sid:
				//System.out.println("LabelSSTRecord类型记录");
				LabelSSTRecord lsrec = (LabelSSTRecord) record;
	
				thisRow = lsrec.getRow();
				thisColumn = lsrec.getColumn();
				if(sstRecord == null) {
					thisStr = '"' + "(No SST Record, can't identify string)" + '"';
				} else {
					thisStr = '"' + sstRecord.getString(lsrec.getSSTIndex()).toString() + '"';
				}
				break;
			case NoteRecord.sid:
				NoteRecord nrec = (NoteRecord) record;
	
				thisRow = nrec.getRow();
				thisColumn = nrec.getColumn();
				// TODO: Find object to match nrec.getShapeId()
				thisStr = '"' + "(TODO)" + '"';
				break;
			case NumberRecord.sid:
				NumberRecord numrec = (NumberRecord) record;
	
				thisRow = numrec.getRow();
				thisColumn = numrec.getColumn();
	
				// Format
				thisStr = formatListener.formatNumberDateCell(numrec);
				break;
			case RKRecord.sid:
				RKRecord rkrec = (RKRecord) record;
	
				thisRow = rkrec.getRow();
				thisColumn = rkrec.getColumn();
				thisStr = '"' + "(TODO)" + '"';
				break;
			default:
				break;
		}
		
		//如下为官方示例代码，为打印输出的细节信息，如：sheet序号、行号
		//不属于内容解析的必要处理，但是如果为了解析过程中，需要生成“x表x行x列”的提示信息时候，可以参考。

		// Handle new row
		if(thisRow != -1 && thisRow != lastRowNumber) {
			lastColumnNumber = -1;
		}

		// Handle missing column
		if(record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord)record;
			thisRow = mc.getRow();
			thisColumn = mc.getColumn();
			thisStr = "";
		}

		
		
		// If we got something to print out, do so
		if(thisStr != null) {
			if(thisColumn > 0) {
				output.print(',');
			}
			output.print(thisStr);
		}

		// Update column and row count
		if(thisRow > -1)
			lastRowNumber = thisRow;
		if(thisColumn > -1)
			lastColumnNumber = thisColumn;

		// Handle end of row
		if(record instanceof LastCellOfRowDummyRecord) {
			// Print out any missing commas if needed
			if(minColumns > 0) {
				// Columns are 0 based
				if(lastColumnNumber == -1) { lastColumnNumber = 0; }
				for(int i=lastColumnNumber; i<(minColumns); i++) {
					output.print(',');
				}
			}

			// We're onto a new row
			lastColumnNumber = -1;

			// End the row
			output.println();
		}
	}
	
	public static void main(String[] args) throws Exception {
		//实例化路径工具
		PathTools pt = new PathTools();
		
		//实例化
		ReadXlsEventModel rem = new ReadXlsEventModel(pt.currentPhysicalPath()+ "jk.xls",-1);
		//执行解析处理
		rem.process();
	}

}
