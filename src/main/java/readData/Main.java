package readData;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class Main {
    public static String dir = "D:\\大创\\";
    public static String fileName = "dataFile.xls";
    public static String[] col = {"二氧化碳(ppm)","甲醛(ug/m3)","TVOC(ug/m3)","PM2.5(ug/m3)","PM10(ug/m3)","温度(℃)","湿度(%RH)","时间"};
    public static List<Map> data = new ArrayList<>();

    public static void main(String[] args) {
        if(!Excel.fileExist(dir + fileName)){
            try {
                Excel.createExcel(dir + fileName,"data",col);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        // 实例化串口操作类对象
        SerialPortUtils serialPort = new SerialPortUtils();
        // 创建串口必要参数接收类并赋值，赋值串口号，波特率，校验位，数据位，停止位
        ParamConfig paramConfig = new ParamConfig("COM7", 9600, 0, 8, 1);
        // 初始化设置,打开串口，开始监听读取串口数据
        serialPort.init(paramConfig);
        // 调用串口操作类的sendComm方法发送数据到串口
        // 关闭串口（注意：如果需要接收串口返回数据的，请不要执行这句，保持串口监听状态）
//        serialPort.closeSerialPort();
        new Timer().schedule(new TimerTask() {
            @Override
            public void run() {
                serialPort.sendComm("0103000000070408");
                ReadDataHex(serialPort.getDataHex());
            }
        },0,1000*5);
    }

    public static void ReadDataHex(String hex){
        System.out.println("=================================");
        if(hex == null){
            System.out.println("暂无数据");
        }else{
            int eCO2 =  Integer.parseInt(hex.substring(6,10),16);//二氧化碳
            int eCH20 = Integer.parseInt(hex.substring(10,14),16);//甲醛
            int TVOC = Integer.parseInt(hex.substring(14,18),16);//TVOC
            int PM2_5 = Integer.parseInt(hex.substring(18,22),16);//PM2.5
            int PM10 = Integer.parseInt(hex.substring(22,26),16);//PM10
            double Temp = Integer.parseInt(hex.substring(26,30),16)*1.0/100;//温度
            double Humi = Integer.parseInt(hex.substring(30,34),16)*1.0/100;//湿度
            System.out.println("二氧化碳：" + eCO2 + "ppm");
            System.out.println("甲醛：" + eCH20 + "ug/m3");
            System.out.println("TVOC：" + TVOC + "ug/m3");
            System.out.println("PM2.5：" + PM2_5 + "ug/m3");
            System.out.println("PM10：" + PM10 + "ug/m3");
            System.out.println("温度：" + Temp + "℃");
            System.out.println("湿度：" + Humi + "%RH");
            Map dataOne = new HashMap();
            dataOne.put("二氧化碳(ppm)",eCO2);
            dataOne.put("甲醛(ug/m3)",eCH20);
            dataOne.put("TVOC(ug/m3)",TVOC);
            dataOne.put("PM2.5(ug/m3)",PM2_5);
            dataOne.put("PM10(ug/m3)",PM10);
            dataOne.put("温度(℃)",Temp);
            dataOne.put("湿度(%RH)",Humi);

            String strDateFormat = "yyyy-MM-dd HH:mm:ss";
            SimpleDateFormat sdf = new SimpleDateFormat(strDateFormat);
            dataOne.put("时间",sdf.format(new Date()));
            System.out.println(sdf.format(new Date()));
            data.add(dataOne);
            try {
                if(data.size() % 10 == 0){
                    Excel.writeToExcel(dir + fileName,"data",data);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
//            public static String[] col = {"二氧化碳(ppm)","甲醛(ug/m3)","TVOC(ug/m3)","PM2.5(ug/m3)","PM10(ug/m3)","温度(℃)","湿度(%RH)"};

        }
    }
    public static void clearScreen() {
        try {
            Runtime.getRuntime().exec("cmd /c cls");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
