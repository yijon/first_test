package yijon;
/**
 * 
 * 路径字符处理工具类
 * 
 * @author yijon 2011-11-15 下午04:51:44
 *
 */
public class PathTools {
	/**
	 * 依据windows或是linux环境，截取出物理路径。
	 * @return
	 */
	public static String subFilePath(String path) {
		// windows路径返回 ：file:/E:/workSpace3.6/SpringTM/WebRoot/WEB-INF/classes/
		// Linux路径返回 : file:/workSpace3.6/SpringTM/WebRoot/WEB-INF/classes/
		String[] paths = path.split(":");
		if (paths.length == 3) {// 如果存在两个":"字符，说明为windows环境，需要去掉"file:/"
			path = path.substring(path.indexOf(":") + 2);
		} else {// 否则就为Linux环境，需要去掉"file:"
			path = path.substring(path.indexOf(":") + 1);
		}
		return path;
	}
	
	/**
	 * 截取出WebRoot物理路径。
	 * @return
	 */
	public static String subWebRootPath(String path) {
		String[] paths = path.split("WEB-INF");
		return paths[0];
	}
	
	/**
	 * 当前java类，所在的物理路径
	 * @return
	 */
	public String currentPhysicalPath() {
		return subFilePath(this.getClass().getResource("").toString());
	}
}
