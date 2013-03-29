package com.lc.xls;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import com.lc.utils.ExcelHandle;
import com.lc.utils.Tool;

import android.app.Activity;
import android.app.AlertDialog;
import android.app.ProgressDialog;
import android.content.Context;
import android.content.DialogInterface;
import android.content.DialogInterface.OnClickListener;
import android.graphics.Typeface;
import android.os.AsyncTask;
import android.os.Bundle;
import android.view.LayoutInflater;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.CheckBox;
import android.widget.CheckedTextView;
import android.widget.CompoundButton;
import android.widget.CompoundButton.OnCheckedChangeListener;
import android.widget.ListView;
import android.widget.TextView;

public class XlsActivity extends Activity {

	private List<Map<String, Object>> mData;
	private ListView listView;
	private Map<Integer, Boolean> checkedStatus;
	private String sumFileName = "";
	private static final String FILE_NAME = "file_name";
	private String path;
	private ProgressDialog mProgressDialog;
	private Typeface tf;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		tf = Typeface.createFromAsset(getAssets(), "font/comic.ttf");
		setContentView(R.layout.activity_xls);
		sumFileName = getString(R.string.sum_file_name);
		checkedStatus = new HashMap<Integer, Boolean>();
		path = Tool.getExcelPath();
		mData = getData();
		ExcelAdapter adapter = new ExcelAdapter(this);
		listView = (ListView) findViewById(R.id.list);
		listView.setAdapter(adapter);
		TextView title = (TextView) findViewById(R.id.title);
		title.setTypeface(tf);
	}

	@Override
	protected void onStart() {
		super.onStart();
		for (int i = 0; i < mData.size(); i++) {// Ä¬ÈÏÈ«Ñ¡
			checkedStatus.put(i, true);
		}
	}

	private List<Map<String, Object>> getData() {
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		List<File> files = Tool.getFiles(new File(Tool.getExcelPath()
				.substring(0, Tool.getExcelPath().length() - 1)));
		if (files != null && files.size() > 0) {
			for (File file : files) {
				if (!sumFileName.equals(file.getName())) {
					Map<String, Object> map = new HashMap<String, Object>();
					map.put(FILE_NAME, file.getName());
					list.add(map);
				}
			}
		}
		return list;
	}

	@Override
	public boolean onCreateOptionsMenu(Menu menu) {
		getMenuInflater().inflate(R.menu.activity_xls, menu);
		return true;
	}

	@Override
	public boolean onMenuItemSelected(int featureId, MenuItem item) {
		switch (item.getItemId()) {
		case R.id.submit:
			final List<String> checkFileNames = new ArrayList<String>();
			final AlertDialog.Builder builder = new AlertDialog.Builder(
					XlsActivity.this);
			final AlertDialog alertDialog = builder.create();
			mProgressDialog = new ProgressDialog(this);
			mProgressDialog.setProgressStyle(ProgressDialog.STYLE_SPINNER);
			mProgressDialog.setTitle(getString(R.string.progress_title));
			mProgressDialog.setMessage(getString(R.string.progress_message));
			mProgressDialog.setCancelable(false);
			mProgressDialog.setButton(DialogInterface.BUTTON_POSITIVE,
					getString(R.string.progress_button),
					new DialogInterface.OnClickListener() {
						@Override
						public void onClick(DialogInterface dialog, int which) {
							mProgressDialog.dismiss();
						}
					});
			Set<Integer> keySet = checkedStatus.keySet();
			for (Integer key : keySet) {
				Map<String, Object> map = mData.get(key);
				String fileName = (String) map.get(FILE_NAME);
				checkFileNames.add(path + fileName);
			}

			final String excelPath = Tool.getExcelPath();
			if (excelPath != null) {
				// files = Tool.getFiles(new File(excelPath));
				new AsyncTask<String, Integer, Integer>() {
					@Override
					protected Integer doInBackground(String... params) {
						List<File> files = new ArrayList<File>();
						for (String fileName : checkFileNames) {
							files.add(new File(fileName));
						}
						return ExcelHandle.sumTotal(files, new File(excelPath
								+ sumFileName));
					}

					@Override
					protected void onPostExecute(Integer result) {
						super.onPostExecute(result);
						if (mProgressDialog.isShowing()) {
							mProgressDialog.dismiss();
						}
						if (result == 1) {
							builder.setTitle(getString(R.string.result_success_title));
							builder.setMessage(getString(R.string.result_success_message));
							builder.setCancelable(true);
						} else {
							builder.setTitle(getString(R.string.result_fault_title));
							builder.setCancelable(true);
						}
						builder.setPositiveButton(
								getString(R.string.result_button),
								new OnClickListener() {
									@Override
									public void onClick(DialogInterface dialog,
											int which) {
										alertDialog.dismiss();
										XlsActivity.this.finish();
									}
								});
						builder.show();
					}
				}.execute();
				mProgressDialog.show();
			} else {
				builder.setTitle(getString(R.string.result_fault_title));
				builder.setCancelable(true);
				builder.setPositiveButton(getString(R.string.result_button),
						new OnClickListener() {
							@Override
							public void onClick(DialogInterface dialog,
									int which) {
								alertDialog.dismiss();
							}
						});
				builder.show();
			}
			break;
		/*case R.id.exit:
			ActivityManager am = (ActivityManager) getSystemService(ACTIVITY_SERVICE);
			am.killBackgroundProcesses(getPackageName());
			break;*/
		default:
			break;
		}

		return super.onMenuItemSelected(featureId, item);
	}
	
	/*@Override
	public void onBackPressed() {
		super.onBackPressed();
		ActivityManager am = (ActivityManager) getSystemService(ACTIVITY_SERVICE);
		am.killBackgroundProcesses(getPackageName());
	}*/

	public class ExcelAdapter extends BaseAdapter {

		private LayoutInflater mInflater;

		public final class ViewHolder {
			public CheckBox checkBox;
			public CheckedTextView checkedTextView;
		}

		public ExcelAdapter(Context context) {
			this.mInflater = LayoutInflater.from(context);
		}

		@Override
		public int getCount() {
			return mData.size();
		}

		@Override
		public Object getItem(int position) {
			return null;
		}

		@Override
		public long getItemId(int position) {
			return 0;
		}

		@Override
		public View getView(final int position, View convertView,
				ViewGroup parent) {
			ViewHolder holder = null;
			if (convertView == null) {
				holder = new ViewHolder();
				convertView = mInflater.inflate(R.layout.item, null);
				holder.checkBox = (CheckBox) convertView
						.findViewById(R.id.check);
				holder.checkedTextView = (CheckedTextView) convertView
						.findViewById(R.id.check_text);
				convertView.setTag(holder);
			} else {
				holder = (ViewHolder) convertView.getTag();
			}
			final String fileName = (String) mData.get(position).get(FILE_NAME);
			holder.checkedTextView.setText(fileName);
			holder.checkedTextView.setTypeface(tf);
			holder.checkBox
					.setOnCheckedChangeListener(new OnCheckedChangeListener() {

						@Override
						public void onCheckedChanged(CompoundButton buttonView,
								boolean isChecked) {
							if (isChecked) {
								checkedStatus.put(position, isChecked);
							} else {
								checkedStatus.remove(position);
							}
						}
					});
			holder.checkBox
					.setChecked(checkedStatus.get(position) == null ? false
							: true);
			return convertView;
		}
	}
}
