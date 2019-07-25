package com.leica.exceldemo;

import android.app.Activity;
import android.os.AsyncTask;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Toast;

import java.util.ArrayList;

import static android.os.Environment.DIRECTORY_DOCUMENTS;

public class MainActivity extends AppCompatActivity implements View.OnClickListener {

    private Activity mActivitySelf = this;
    private static final String TAG = "MainActivity";
    POIAndroid mPOIAndroid = new POIAndroid();
    Button mButtonWrite;
    Button mButtonRead;
    //写入的文件(包括路径与文件名)
    String file;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        InitControl();
        mButtonRead.setOnClickListener(this);
        mButtonWrite.setOnClickListener(this);
        file = getExternalFilesDir(DIRECTORY_DOCUMENTS).getPath() + "/" + "content.xlsx";
    }

    private void InitControl() {
        mButtonRead = findViewById(R.id.btn_read);
        mButtonWrite = findViewById(R.id.btn_write);
    }

    @Override
    public void onClick(View v) {
        switch (v.getId()) {
            case R.id.btn_read:
                ReadTask readTask = new ReadTask();
                readTask.execute(file);
                break;
            case R.id.btn_write:
                WriteTask writeTask = new WriteTask();
                String temp = ((EditText) findViewById(R.id.et_write)).getText().toString();
                writeTask.execute(temp);
                break;
            default:
                Log.e(TAG, "onClick: " + "ERROR");
        }
    }

    private class ReadTask extends AsyncTask<String, Long, String> {

        @Override
        protected String doInBackground(String... strings) {
            ArrayList<Record> records = mPOIAndroid.readXLSXFile(strings[0]);
            return null == records ? null : records.get(0).getName();
        }

        @Override
        protected void onPostExecute(String s) {
            super.onPostExecute(s);
            if (null == s) {
                Toast.makeText(mActivitySelf, "读取失败,请先点击写入按钮", Toast.LENGTH_SHORT).show();
            } else {
                ((EditText) findViewById(R.id.et_read)).setText(s);
                Toast.makeText(mActivitySelf, "读取成功", Toast.LENGTH_SHORT).show();
            }
        }
    }

    private class WriteTask extends AsyncTask<String, Long, String> {

        @Override
        protected String doInBackground(String... strings) {

            try {
                Record record = new Record();
                record.setName(strings[0]);
                boolean bFileExist = FileUtil.isFileExist(file);
                if (bFileExist) {
                    mPOIAndroid.insertRows(file, record);
                } else {
                    mPOIAndroid.create(file, record);
                }
                return "ture";

            } catch (Exception ex) {
                return null;
            }
        }

        @Override
        protected void onPostExecute(String s) {
            super.onPostExecute(s);
            if (null == s) {
                Toast.makeText(mActivitySelf, "写入失败", Toast.LENGTH_SHORT).show();
            } else {
                Toast.makeText(mActivitySelf, "写入成功", Toast.LENGTH_SHORT).show();
            }
        }
    }
}
