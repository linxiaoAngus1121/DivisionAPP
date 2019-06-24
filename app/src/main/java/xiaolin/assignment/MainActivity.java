package xiaolin.assignment;

import android.Manifest;
import android.annotation.SuppressLint;
import android.app.ActionBar;
import android.app.AlertDialog;
import android.app.Application;
import android.app.ProgressDialog;
import android.content.ActivityNotFoundException;
import android.content.ContentUris;
import android.content.Context;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.content.res.AssetManager;
import android.database.Cursor;
import android.net.Uri;
import android.os.Build;
import android.os.Environment;
import android.os.Handler;
import android.os.Message;
import android.provider.DocumentsContract;
import android.provider.MediaStore;
import android.support.annotation.NonNull;
import android.support.v4.app.ActivityCompat;
import android.support.v4.content.ContextCompat;
import android.support.v4.content.FileProvider;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.text.TextUtils;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;

import com.geek.thread.GeekThreadManager;
import com.geek.thread.ThreadPriority;
import com.geek.thread.ThreadType;
import com.geek.thread.task.GeekRunnable;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import xiaolin.assignment.bean.UserBean;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    private static final String TAG = "000";

    private List<UserBean> list = new ArrayList<>();    //总的人数，包含头部的性别，学号，姓名
    private List<UserBean> maleList = new ArrayList<>();        //男
    private List<UserBean> falemanList = new ArrayList<>();        //女
    private Handler handler;

    //xlsx文件解析
    private static final int PARSEERRORCODE = 0x11;
    private static final int PARSESUCCESSCODE = 0x12;

    //xlsx文件生成
    private static final int CREATEERRORCODE = 0x13;
    private static final int CREATESUCCESSCODE = 0x14;


    //分班失败
    private static final int PARSEERROR = 0x15;

    //提示信息
    private static final int TIPONE = 0x16;
    private static final int TIPTWO = 0x17;
    private static final int TIPSECOND = 0x21;

    //权限
    private static final int MY_PERMISSIONS_REQUEST_EXTERNAL_STORAGE = 0x18;

    //显示分班dialog
    private static final int SHOW_DIALOG = 0x19;
    private static final int HIDE_DIALOG = 0x20;

    private TextView mTvParse;
    private TextView mTvDivide;
    private TextView mTvpath;
    private EditText mEtban;
    private ProgressDialog dialog;
    private ProgressDialog pd2;
    private Button mBtclear;
    private Button mBtchoise;
    private boolean ok = false;
    private AlertDialog alertDialog;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        initView();
        createFloder();

        alertDialog = new AlertDialog.Builder(this)
                .setPositiveButton("我知道了", null)
                .setMessage("请确保文件管理器根目录中有名为‘工作簿.xlsx’文件或手动选择文件。分班完成后文件将保存到根目录的分班文件夹中")
                .create();
        alertDialog.show();
    }

    private void createFloder() {
        File cacheDir = Environment.getExternalStorageDirectory();
        String s = cacheDir.getPath() + "/分班/";
        File outFile = new File(s);
        if (!outFile.exists()) {
            ok = outFile.mkdirs();
        } else {
            ok = true;
        }
    }

    @SuppressLint("HandlerLeak")
    private void initView() {
        mTvParse = findViewById(R.id.tv_parse);
        mTvDivide = findViewById(R.id.tv_divide);
        mEtban = findViewById(R.id.et_ban);
        mTvpath = findViewById(R.id.tv_path);
        mBtclear = findViewById(R.id.bt_clearTv);
        mBtchoise = findViewById(R.id.chioseFile);
        mTvParse.setOnClickListener(this);
        mTvDivide.setOnClickListener(this);
        mBtclear.setOnClickListener(this);
        mBtchoise.setOnClickListener(this);
        dialog = new ProgressDialog(this);
        dialog.setTitle("excel文件解析");
        dialog.setMessage("excel文件正在解析，解析进度取决于文件大小，请耐心等待...");
        dialog.setCancelable(false);
        dialog.setProgressStyle(ProgressDialog.STYLE_SPINNER);


        pd2 = new ProgressDialog(MainActivity.this);     //分班进度条
        pd2.setTitle("正在分班");
        pd2.setCancelable(false);
        pd2.setProgressStyle(ProgressDialog.STYLE_HORIZONTAL);
        pd2.setMax(100);
        handler = new Handler() {
            @Override
            public void handleMessage(Message msg) {
                super.handleMessage(msg);
                dialog.cancel();
                switch (msg.what) {
                    case PARSEERRORCODE:
                        Toast.makeText(MainActivity.this, "打开文件失败,请检查文件是否放置在根目录", Toast.LENGTH_SHORT).show();
                        break;
                    case PARSESUCCESSCODE:
                        Toast.makeText(MainActivity.this, "文件解析成功,一共" + list.size() + "条数据", Toast.LENGTH_SHORT).show();
                        break;
                    case CREATEERRORCODE:
                        Toast.makeText(MainActivity.this, "第" + msg.arg1 + "分班文件创建失败", Toast.LENGTH_SHORT).show();
                        break;
                    case CREATESUCCESSCODE:
                        pd2.setProgress(msg.arg1);
                        break;
                    case PARSEERROR:
                        Toast.makeText(MainActivity.this, "分班信息出错", Toast.LENGTH_SHORT).show();
                        break;
                    case TIPONE:
                        Toast.makeText(MainActivity.this, "excel文件解析出错或者未解析excel", Toast.LENGTH_SHORT).show();
                        break;
                    case TIPTWO:
                        Toast.makeText(MainActivity.this, "请输入正确的班级个数", Toast.LENGTH_SHORT).show();
                        break;
                    case SHOW_DIALOG:
                        pd2.show();
                        break;
                    case HIDE_DIALOG:
                        /*  Toast.makeText(MainActivity.this, "所有文件生成完毕", Toast.LENGTH_SHORT).show();
                      String path = Environment.getExternalStorageDirectory().getAbsolutePath() + "/分班/filetoshare-1.xlsx";
                        File file = new File(path);
                        Log.i(TAG, "handleMessage: " + path);
                        if (!file.exists()) {
                            return;
                        }
                        Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                        intent.addCategory(Intent.CATEGORY_DEFAULT);
                        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
                        Uri uri = FileProvider.getUriForFile(MainActivity.this, "xiaolin.assignment.fileprovider", file);
                        intent.setDataAndType(uri, "file/*");
                        try {
                            startActivity(intent);
                        } catch (ActivityNotFoundException e) {
                            e.printStackTrace();
                        }*/
                        alertDialog.setTitle("温馨提示");
                        if (ok) {
                            alertDialog.setMessage("文件已保存到根目录中的分班文件夹中，请查看");
                        } else {
                            alertDialog.setMessage("文件已保存到根目录中，请查看");
                        }
                        alertDialog.show();
                        pd2.cancel();
                        break;
                    case TIPSECOND:
                        Toast.makeText(MainActivity.this, "只能识别以xlsx格式的excel文件", Toast.LENGTH_SHORT).show();
                        break;

                }
            }
        };
    }


    private void deal(int c, UserBean bean, String value) {
        switch (c % 4) {
            case 0:
                bean.setNum(value);
                break;
            case 1:
                bean.setName(value);
                break;
            case 2:
                bean.setSex(value);
                break;
            case 3:
                bean.setIdcard(value);
                break;
        }
    }


    /**
     * 解析excel
     */
    public void parse() {
        if (!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)) {
            Toast.makeText(this, "手机未安装sd卡", Toast.LENGTH_SHORT).show();
            return;
        }
        dialog.show();
        clearData();
        GeekThreadManager.getInstance().execute(new GeekRunnable(ThreadPriority.NORMAL) {
            @Override
            public void run() {
                try {
                    String defaultPath = mTvpath.getText().toString();
                    if (TextUtils.isEmpty(defaultPath)) {     //代表他没有选择文件，那就去根目录找
                        Log.i(TAG, "run: 使用默认的");
                        defaultPath = Environment.getExternalStorageDirectory().getPath() + "/工作簿.xlsx";
                        File file = new File(defaultPath);
                        if (!file.exists()) {
                            handler.sendEmptyMessage(PARSEERRORCODE);
                            return;
                        }
                    }
                    if (!defaultPath.endsWith(".xlsx")) {       //检查后缀
                        handler.sendEmptyMessage(TIPSECOND);
                        return;
                    }
                    XSSFWorkbook workbook = new XSSFWorkbook(defaultPath);              //整个xLs文件
                    XSSFSheet sheet = workbook.getSheetAt(0);               //代表一张表，即sheet1
                    int rowsCount = sheet.getPhysicalNumberOfRows();                //总共有多少行
                    Log.i(TAG, "parse: " + rowsCount);
                    FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    Row sheetRow = sheet.getRow(1);
                    int cellsCount = sheetRow.getPhysicalNumberOfCells();       //总共多少列
                    int time = cellsCount / 4;              //几列，/4 就是循环的次数
                    Log.i(TAG, "parse:次数 " + time);
                    int t = 0;
                    for (int q = 0; q < time; q++) {                 //循环列的7次
                        for (int r = 0; r < rowsCount; r++) {       //循环行的
                            if (r == 0 || r == 1) {
                                continue;
                            }
                            UserBean bean = new UserBean();
                            Row row = sheet.getRow(r);
                            for (int i = t; i < t + 4; i++) {       //取4个
                                String value = getCellAsString(row, i, formulaEvaluator);
                                if (value == null) {
                                    bean.setAdd(0);
                                    continue;
                                }
                                deal(i, bean, value);
                            }
                            if (bean.getAdd() == 1) {
                                list.add(bean);
                            }
                            //对男女进行分组
                            if (TextUtils.equals(bean.getSex(), "男")) {
                                maleList.add(bean);
                            } else if (TextUtils.equals(bean.getSex(), "女")) {
                                falemanList.add(bean);
                            }
                        }
                        t = t + 4;
                    }

                    Log.i(TAG, "parse: list" + list.size());
                    handler.sendEmptyMessage(PARSESUCCESSCODE);
                } catch (IOException e) {
                    e.printStackTrace();
                    Log.i(TAG, "parse: 打开文件失败");
                    handler.sendEmptyMessage(PARSEERRORCODE);
                }
            }
        }, ThreadType.NORMAL_THREAD);
    }

    /**
     * 分班
     */
    public void assignment() {
        final int size = list.size();
        final int mansize = maleList.size();
        final int femalsize = falemanList.size();
        if (size == 0 || mansize == 0 || femalsize == 0) {
            handler.sendEmptyMessage(TIPONE);
            Log.i(TAG, "assignment: 有人是0");
            return;
        }
        //正式分班
        GeekThreadManager.getInstance().execute(new GeekRunnable(ThreadPriority.NORMAL) {
            @Override
            public void run() {
                try {

                    int ban = 0;  //分几个班
                    String banst = mEtban.getText().toString();
                    if (TextUtils.isEmpty(banst)) {
                        ban = 7;
                    } else {
                        if (!TextUtils.isDigitsOnly(banst)) {
                            handler.sendEmptyMessage(TIPTWO);
                            return;
                        }
                        ban = Integer.valueOf(banst);
                    }

                    int manduo = mansize % ban;     //男的多出来几个

                    int i1 = mansize / ban;         //男的一个班有多少人

                    int maleduo = femalsize % ban;      //女的多出来几个

                    int i2 = femalsize / ban;      //女的一个班有多少人

                    // 男的第一次0-i1  女的第一次 0-i2 分别取ban次

                    ArrayList<ArrayList<UserBean>> lists = new ArrayList<>(ban);        //分几个班，就几个list
                    int t = 0;
                    int q = 0;
                    for (int j = 0; j < ban; j++) {

                        ArrayList<UserBean> list = new ArrayList<>();

                        //男的先取
                        for (int k = t; k <= t + i1 - 1; k++) {
                            list.add(maleList.get(k));
                        }
                        t = t + i1;

                        //女的取
                        for (int v = q; v <= q + i2 - 1; v++) {
                            list.add(falemanList.get(v));
                        }
                        q = q + i2;
                        lists.add(list);
                    }

                    //处理余数

                    //有余数的情况
                    if (maleduo != 0 || manduo != 0) {
                        while (t < mansize || q < femalsize) {

                            for (int i3 = 0; i3 < ban; i3++) {      //几个班一起循环，然后逐个班添加进去
                                if (t == mansize && q == femalsize) {
                                    break;
                                }
                                if (t < mansize) {  //男的添加完了就走else
                                    lists.get(i3).add(maleList.get(t));
                                    t++;
                                } else {
                                    if (q < femalsize) {
                                        lists.get(i3).add(falemanList.get(q));
                                        q++;
                                    }
                                }

                            }
                        }
                    }
                    Log.i(TAG, "assignment: lists的size" + lists.size());

                    if (lists.size() == 0) {
                        handler.sendEmptyMessage(PARSEERROR);
                    } else {
                        //生成分班后的各个文件
                        handler.sendEmptyMessage(SHOW_DIALOG);
                        for (int j = 0; j < lists.size(); j++) {
                            create(lists.get(j), j + 1, ban);
                        }
                        handler.sendEmptyMessage(HIDE_DIALOG);
                    }
                    clearData();
                } catch (Exception e) {
                    e.printStackTrace();
                    handler.sendEmptyMessage(PARSEERROR);
                }
            }
        }, ThreadType.NORMAL_THREAD);
    }

    /**
     * 清除解析完成后的数据
     */
    private void clearData() {
        list.clear();
        maleList.clear();
        falemanList.clear();
    }

    /**
     * 将分班信息写成每一个excel
     *
     * @param bean 班级信息
     * @param i    第几个
     * @param ban
     */
    public void create(ArrayList<UserBean> bean, int i, int ban) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName("第" + i + "个班"));
            XSSFRow row1 = sheet.createRow(0);
            for (int j = 0; j < 4; j++) {
                XSSFCell cell = row1.createCell(j);
                if (j == 0) {       //添加上面的说明
                    cell.setCellValue("序号");
                } else if (j == 1) {
                    cell.setCellValue("名单");
                } else if (j == 2) {
                    cell.setCellValue("性别");
                } else {
                    cell.setCellValue("身份证后四位");
                }
            }
            for (int j = 0; j < bean.size(); j++) {     //将list中的数据全部取出
                Row row = sheet.createRow(j + 1);
                UserBean bean1 = bean.get(j);
                for (int a = 0; a < 4; a++) {
                    Cell cell = row.createCell(a);
                    if (a == 0) {
                        cell.setCellValue(bean1.getNum());
                    } else if (a == 1) {
                        cell.setCellValue(bean1.getName());
                    } else if (a == 2) {
                        cell.setCellValue(bean1.getSex());
                    } else {
                        cell.setCellValue(bean1.getIdcard());
                    }
                }
            }
            String outFileName = "filetoshare-" + i + ".xlsx";
            File cacheDir = Environment.getExternalStorageDirectory();
            File outFile;
            if (ok) {     //创建在分班文件夹里面
                outFile = new File(cacheDir.getAbsolutePath() + "/分班/", outFileName);
            } else {        //创建在根目录
                outFile = new File(cacheDir, outFileName);
            }
            OutputStream outputStream = new FileOutputStream(outFile.getAbsolutePath());
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            Message message = handler.obtainMessage();
            message.what = CREATESUCCESSCODE;
            double pressent = (double) i / ban * 100;
            int progress = (int) Math.ceil(pressent);
            message.arg1 = progress;
            handler.sendMessage(message);
        } catch (Exception e) {
            printlnToUser(e.toString());
            Message message = handler.obtainMessage();
            message.what = CREATEERRORCODE;
            message.arg1 = i;
            handler.sendMessage(message);
        }
    }


    private void printlnToUser(String cellInfo) {
        Log.i(TAG, "printlnToUser: cellInfo" + cellInfo);
    }


    protected String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
        String value = "";
        try {
            Cell cell = row.getCell(c);     //cell代表一个单元格
            CellValue cellValue = formulaEvaluator.evaluate(cell);
            switch (cellValue.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                    value = "" + cellValue.getBooleanValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    int numericValue = (int) cellValue.getNumberValue();
                    value = String.valueOf(numericValue);
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = "" + cellValue.getStringValue();
                    break;
                default:
            }
        } catch (NullPointerException e) {
            printlnToUser(e.toString());
            return null;
        }
        return value;
    }


    @Override
    public void onClick(View view) {
        switch (view.getId()) {
            case R.id.tv_parse:     //解析excel
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
                    if (ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
                        ActivityCompat.requestPermissions(this,
                                new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE},
                                MY_PERMISSIONS_REQUEST_EXTERNAL_STORAGE);
                    } else {
                        parse();
                    }
                } else {
                    parse();
                }
                break;
            case R.id.tv_divide:    //分班
                assignment();
                break;
            case R.id.chioseFile:
                choise();
                break;
            case R.id.bt_clearTv:
                clear();
                break;
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
        if (requestCode == MY_PERMISSIONS_REQUEST_EXTERNAL_STORAGE) {
            if (grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                parse();
            } else {
                Toast.makeText(this, "权限被拒绝", Toast.LENGTH_SHORT).show();
                finish();
            }
        }
    }

    public void choise() {
        Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
        intent.setType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");    //只选择xlsx后缀的文件
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        startActivityForResult(intent, 2);
    }


    String path1;

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode == 2 && resultCode == RESULT_OK) {
            Uri uri = data.getData();//得到uri，后面就是将uri转化成file的过程。
            if ("file".equalsIgnoreCase(uri.getScheme())) {//使用第三方应用打开
                path1 = uri.getPath();
                return;
            }
            path1 = getPath(this, uri);
            mTvpath.setText(path1);
        }

    }

    private String getPath(Context context, Uri uri) {


        // DocumentProvider
        if (DocumentsContract.isDocumentUri(context, uri)) {
            // ExternalStorageProvider
            if (isExternalStorageDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];


                if ("primary".equalsIgnoreCase(type)) {
                    return Environment.getExternalStorageDirectory() + "/" + split[1];
                }
            }
            // DownloadsProvider
            else if (isDownloadsDocument(uri)) {


                final String id = DocumentsContract.getDocumentId(uri);
                final Uri contentUri = ContentUris.withAppendedId(
                        Uri.parse("content://downloads/public_downloads"), Long.valueOf(id));


                return getDataColumn(context, contentUri, null, null);
            }
            // MediaProvider
            else if (isMediaDocument(uri)) {
                final String docId = DocumentsContract.getDocumentId(uri);
                final String[] split = docId.split(":");
                final String type = split[0];


                Uri contentUri = null;
                if ("image".equals(type)) {
                    contentUri = MediaStore.Images.Media.EXTERNAL_CONTENT_URI;
                } else if ("video".equals(type)) {
                    contentUri = MediaStore.Video.Media.EXTERNAL_CONTENT_URI;
                } else if ("audio".equals(type)) {
                    contentUri = MediaStore.Audio.Media.EXTERNAL_CONTENT_URI;
                }


                final String selection = "_id=?";
                final String[] selectionArgs = new String[]{split[1]};


                return getDataColumn(context, contentUri, selection, selectionArgs);
            }
        }
        // MediaStore (and general)
        else if ("content".equalsIgnoreCase(uri.getScheme())) {
            return getDataColumn(context, uri, null, null);
        }
        // File
        else if ("file".equalsIgnoreCase(uri.getScheme())) {
            return uri.getPath();
        }
        return null;
    }

    public boolean isExternalStorageDocument(Uri uri) {
        return "com.android.externalstorage.documents".equals(uri.getAuthority());
    }

    public boolean isMediaDocument(Uri uri) {
        return "com.android.providers.media.documents".equals(uri.getAuthority());
    }

    public boolean isDownloadsDocument(Uri uri) {
        return "com.android.providers.downloads.documents".equals(uri.getAuthority());
    }

    public String getDataColumn(Context context, Uri uri, String selection,
                                String[] selectionArgs) {


        Cursor cursor = null;
        final String column = "_data";
        final String[] projection = {column};


        try {
            cursor = context.getContentResolver().query(uri, projection, selection, selectionArgs,
                    null);
            if (cursor != null && cursor.moveToFirst()) {
                final int column_index = cursor.getColumnIndexOrThrow(column);
                return cursor.getString(column_index);
            }
        } finally {
            if (cursor != null)
                cursor.close();
        }
        return null;
    }

    public void clear() {
        mTvpath.setText("");
    }
}
