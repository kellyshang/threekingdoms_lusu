import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.ArrayList;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ConvertXmindFile {

  static String subjectMain;
  static ArrayList<String> secondSubjectsArrayList = new ArrayList<String>(); //二级主题数组
  static ArrayList<String> thirdSubjectsArrayList = new ArrayList<String>();
  static ArrayList<String> fourthSubjectsArrayList = new ArrayList<String>();
  static ArrayList<String> fifthSubjectsArrayList = new ArrayList<String>();
  static ArrayList<String> sixthSubjectsArrayList = new ArrayList<String>();


  public static String readJsonFile(String fileName) {
    String jsonStr = "";
    try {
      File jsonFile = new File(fileName);
      FileReader fileReader = new FileReader(jsonFile);
      Reader reader = new InputStreamReader(new FileInputStream(jsonFile), "utf-8");
      int ch = 0;
      StringBuffer sb = new StringBuffer();
      while ((ch = reader.read()) != -1) {
        sb.append((char) ch);
      }
      fileReader.close();
      reader.close();
      sb.deleteCharAt(0);
      sb.deleteCharAt(sb.length() - 1);
      jsonStr = sb.toString();
      return jsonStr;
    } catch (IOException e) {
      e.printStackTrace();
      return null;
    }
  }

  /**
   * @return void
   * @author kellyshang
   * @date 2020/10/6
   * @description 读取子元素，并相应的给父级数组添加相应数量的空值
   */
  static void readChildren(JSONArray jsonArray, JSONArray childJsonArray,
      int columnNum) {
    int tmpSize = 0;
    // 循环三级数组，如果children不为空，则存在四级数组，然后获取attached数组；若为空，则在四级数组里填入空值；
    for (int k = 0; k < jsonArray.size(); k++) {
      if ("".equals(jsonArray.get(k))
          || jsonArray.getJSONObject(k).getJSONObject("children") == null) {
        childJsonArray.add("");
        switch (columnNum) {
          case 4:
            fourthSubjectsArrayList.add("");
            break;
          case 5:
            fifthSubjectsArrayList.add("");
            break;
          case 6:
            sixthSubjectsArrayList.add("");
            break;
          default:
            throw new IllegalArgumentException("The argument is not correct");
        }
        tmpSize += 1;
        System.out.printf("~~~k & tmpsize:~~ %d, %d。\n", k, tmpSize);
      } else {
        JSONArray childrenJsonArray4 = jsonArray.getJSONObject(k).getJSONObject("children")
            .getJSONArray("attached");

        // 判断四级主题数量，大于1，则三级主题数组增加(n-1)个空值
        if (childrenJsonArray4.size() > 1) {
          for (int i = 0; i < childrenJsonArray4.size() - 1; i++) {
            System.out.println(k);
            // tmpSize取之前的所有最低子元素之和，再加上当前元素+1，即在当前元素之后的位置开始加空值
            // 当前元素加了空值后，当前的父元素也需要层层迭加空值
            switch (columnNum) {
              case 4:
                thirdSubjectsArrayList.add(tmpSize + 1, "");
                secondSubjectsArrayList.add(tmpSize + 1, "");
                break;
              case 5:
                fourthSubjectsArrayList.add(tmpSize + 1, "");
                thirdSubjectsArrayList.add(tmpSize + 1, "");
                secondSubjectsArrayList.add(tmpSize + 1, "");
                break;
              case 6:
                fifthSubjectsArrayList.add(tmpSize + 1, "");
                fourthSubjectsArrayList.add(tmpSize + 1, "");
                thirdSubjectsArrayList.add(tmpSize + 1, "");
                secondSubjectsArrayList.add(tmpSize + 1, "");
                break;
              default:
                throw new IllegalArgumentException("The argument is not correct");
            }

          }
        }
        for (int l = 0; l < childrenJsonArray4.size(); l++) {
          childJsonArray.add(childrenJsonArray4.getJSONObject(l));
          switch (columnNum) {
            case 4:
              fourthSubjectsArrayList.add(childrenJsonArray4.getJSONObject(l).getString("title"));
              break;
            case 5:
              fifthSubjectsArrayList.add(childrenJsonArray4.getJSONObject(l).getString("title"));
              break;
            case 6:
              sixthSubjectsArrayList.add(childrenJsonArray4.getJSONObject(l).getString("title"));
              break;
            default:
              throw new IllegalArgumentException("The argument is not correct");
          }
        }
        tmpSize += childrenJsonArray4.size();
        System.out.printf("k & tmpsize: %d, %d。\n", k, tmpSize);
      }
    }
  }

  static void writeToExcel(String excelPath) throws IOException, WriteException {
    File file = new File(excelPath);
    file.createNewFile();
    WritableWorkbook workbook = Workbook.createWorkbook(file);
    WritableSheet sheet = workbook.createSheet("sheet1", 0);
    String[] title = {"TestSuite", "Module", "Cases",};
    Label label = null;

    for (int i = 0; i < title.length; i++) {
      label = new Label(i, 0, title[i]);
      sheet.addCell(label);
    }

    // 第1列第2行设置主题名称
    label = new Label(1, 1, subjectMain);
    sheet.addCell(label);

    // 从第2列第2行开始追加
    for (int i = 1; i <= secondSubjectsArrayList.size(); i++) {
      label = new Label(1, i, secondSubjectsArrayList.get(i - 1));
      sheet.addCell(label);
    }

    // 从第3列第2行开始追加
    for (int i = 1; i <= thirdSubjectsArrayList.size(); i++) {
      label = new Label(2, i, thirdSubjectsArrayList.get(i - 1));
      sheet.addCell(label);
    }

    // 从第4列第2行开始追加
    for (int i = 1; i <= fourthSubjectsArrayList.size(); i++) {
      label = new Label(3, i, fourthSubjectsArrayList.get(i - 1));
      sheet.addCell(label);
    }

    // 从第5列第2行开始追加
    for (int i = 1; i <= fifthSubjectsArrayList.size(); i++) {
      label = new Label(4, i, fifthSubjectsArrayList.get(i - 1));
      sheet.addCell(label);
    }

    // 从第6列第2行开始追加
    for (int i = 1; i <= sixthSubjectsArrayList.size(); i++) {
      label = new Label(5, i, sixthSubjectsArrayList.get(i - 1));
      sheet.addCell(label);
    }

    workbook.write();
    workbook.close();
  }


  public static void main(String[] args) throws IOException, WriteException {
    String xmindJsonString = readJsonFile(
        System.getProperty("user.dir") + "/xmindFile/content.json");
    JSONObject jsonObject = JSONObject.parseObject(xmindJsonString);
    subjectMain = jsonObject.getJSONObject("rootTopic").getString("title");
    System.out.println(subjectMain); // 获取一级主题
    System.out.println(jsonObject.getJSONObject("rootTopic").getJSONObject("children"));
    // 二级主题
    JSONArray secondJsonArray = jsonObject.getJSONObject("rootTopic").getJSONObject("children")
        .getJSONArray("attached");
    System.out.println("二级主题数量" + secondJsonArray.size());

    JSONArray thirdJsonArray = new JSONArray();
    JSONArray fourthJsonArray = new JSONArray();
    JSONArray fifthJsonArray = new JSONArray();
    JSONArray sixthJsonArray = new JSONArray();

    for (int i = 0; i < secondJsonArray.size(); i++) {
      secondSubjectsArrayList.add(secondJsonArray.getJSONObject(i).getString("title"));
      System.out.println("****" + secondSubjectsArrayList.get(i)); //打印二级主题

      // 三级主题(必不为空)
      JSONArray childrenJsonArray3 = secondJsonArray.getJSONObject(i).getJSONObject("children")
          .getJSONArray("attached");
      // 判断三级主题数量，如果大于1，则二级主题数组相应增加(n-1)个空值
      if (childrenJsonArray3.size() > 1) {
        for (int j = 0; j < childrenJsonArray3.size() - 1; j++) {
          secondSubjectsArrayList.add("");
        }
      }
      for (int j = 0; j < childrenJsonArray3.size(); j++) {
        thirdJsonArray.add(childrenJsonArray3.getJSONObject(j));
        thirdSubjectsArrayList
            .add(childrenJsonArray3.getJSONObject(j).getString("title"));
      }
    }

    readChildren(thirdJsonArray, fourthJsonArray, 4);
    readChildren(fourthJsonArray, fifthJsonArray, 5);
    readChildren(fifthJsonArray, sixthJsonArray, 6);

    System.out.println("~~~~~~~三级主题：~~~~~~~~~~");
    for (int i = 0; i < thirdSubjectsArrayList.size(); i++) {
      System.out.println(i + thirdSubjectsArrayList.get(i));
    }

    System.out.println("~~~~~~~四级主题：~~~~~~~~~~");
    for (int i = 0; i < fourthSubjectsArrayList.size(); i++) {
      System.out.println(i + fourthSubjectsArrayList.get(i));
    }

    System.out.println("~~~~~~~五级主题：~~~~~~~~~~");
    for (int i = 0; i < fifthSubjectsArrayList.size(); i++) {
      System.out.println(fifthSubjectsArrayList.get(i));
    }

    System.out.println("~~~~~~~六级主题：~~~~~~~~~~");
    for (int i = 0; i < sixthSubjectsArrayList.size(); i++) {
      System.out.println(sixthSubjectsArrayList.get(i));
    }

    // write to excel
    writeToExcel(System.getProperty("user.dir") + "/xmind_cases.xlsx");
  }

}
