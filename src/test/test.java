package test;

import java.net.InetAddress;
import java.net.UnknownHostException;

public class test {

	public static void main(String[] args) throws UnknownHostException 
	{
		System.out.println(System.getProperty("user.name"));
		InetAddress inetAddress = InetAddress.getLocalHost();
		System.out.println(inetAddress.getHostAddress());

		System.out.println(System.getProperty("os.name"));
	}

}
