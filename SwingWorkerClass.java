package firstdemo;

import java.awt.event.WindowEvent;

import javax.swing.JFrame;
import javax.swing.JProgressBar;
import javax.swing.SwingWorker;

public class SwingWorkerClass extends SwingWorker{
	private static String[] User = null;
	private static  Boolean[] Result = null;
	private static  String[] Env = null;
	private static int progress=0;
	JFrame F;
	JProgressBar p;
	public SwingWorkerClass(String[] env, String[] user, Boolean[] result) {
		// TODO Auto-generated method stub
		this.execute();
		User=user;
		Env=env;
		Result=result;
		F=new JFrame("Progress");
		F.setSize(400, 200);
		p=new JProgressBar();
		p.setBounds(10, 10, 380, 20);
		F.add(p);
		F.setVisible(true);
		F.setDefaultCloseOperation(F.EXIT_ON_CLOSE);
	}
	@Override
	protected Object doInBackground() throws Exception {
		// TODO Auto-generated method stub
		for(int i=0;i<6;i++) {
			Excel.SendResult(User[i], Env[i], Result[i]);
			Thread.sleep(2000);
			p.setValue(progress*100/6);
			progress++;
			
		}
		//Excel.FillSheet();
		return null;
	}
	protected void done() {  
		F.dispatchEvent(new WindowEvent(F, WindowEvent.WINDOW_CLOSING));
		
    } 
	

	

}
