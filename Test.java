package firstdemo;

import java.awt.HeadlessException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;

public class Test {
	static JFrame f;
	static JButton start;
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		final String[] User = {"user1","user4","user2","user5","user6","user7"};
		final String[] Env = {"Env1","Env3","Env4","Env6","Env9","Env2"};
		final Boolean[] Result = {true,false,true,true,true,false};
		try {
			f=new JFrame("Main");
			f.setSize(200,200);
			start=new JButton("Start");
			start.setBounds(10, 10, 100, 20);
			f.add(start);
			f.setVisible(true);
			
		} catch (HeadlessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		start.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				SwingWorkerClass obj=new SwingWorkerClass(Env,User,Result);
			}
		});
		//Excel.SendResult("user1","Env2",false);
		
	}
		

}
