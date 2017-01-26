
public class KillApp extends Thread {
	
	private Thread searchEngineThread;
	
	KillApp(Thread searchEngineThread) {
		this.searchEngineThread = searchEngineThread;
		
	}
	
	public void run() {
		
		
		while (searchEngineThread.isAlive()) {
			//System.out.println("Kill App: waiting search engine to die...");
			
			try {
			
				sleep(500);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				//System.out.println("Kill App: interrupted exception");
				//e.printStackTrace();
			}
			
		}
		
		System.exit(0);
		
	}

	
	
	
}
