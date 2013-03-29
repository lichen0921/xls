package com.lc.xls;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;

public class ExitActivity extends Activity {

	public static boolean flag = false;
	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		if (flag) {
			flag = false;
			finish();
		} else {
			Intent intent = new Intent(this, XlsActivity.class);
			startActivity(intent);
		}
	}
}
