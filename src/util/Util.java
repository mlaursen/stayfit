package util;

import java.util.ArrayList;
import java.util.List;

public class Util {

	public static int[] numsFrom(int low, int high) {
		int[] list = new int[high + 1];
		for(int i = low; i <= high; i++)
			list[i] = i;
		return list;
	}
	
	
}
