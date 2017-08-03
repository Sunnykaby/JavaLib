package Thread;

/**
 * Created by User on 2017/8/2.
 */
public class ThreadSafeTest {
    public static void main(String[] args) throws InterruptedException{
        ProcessThread pt = new ProcessThread();
        Thread t1 = new Thread(pt, "t1");
        t1.start();
        Thread.sleep(10000);
        Thread t2 = new Thread(pt, "t2");
        t2.start();
        t1.join();
        t2.join();
        System.out.println(pt.getCount());
    }
}

class ProcessThread implements Runnable {
    private int count = 0;

    @Override
    public void run() {
        for (int i = 0; i < 100; i++) {
            sleepThread(i);
            count++;//Here is the problem
            System.out.println(Thread.currentThread().getName() + ":" + count);
        }
    }

    public int getCount() {
        return count;
    }

    public void sleepThread(int i) {
        try {
            Thread.sleep(i * 100);
        } catch (InterruptedException e) {
            System.out.println(e);
        }
    }
}
