package org.sang.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.sang.bean.*;
import org.sang.mapper.*;
import org.sang.utils.ExcelUtil;
import org.sang.utils.Util;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by sang on 2017/12/17.
 */
@Service
@Transactional
public class UserService implements UserDetailsService {
    @Autowired
    UserMapper userMapper;
    @Autowired
    RolesMapper rolesMapper;
    @Autowired
    PasswordEncoder passwordEncoder;

    @Autowired
    PersonManageMapper personManageMapper;

    @Autowired
    OutWorkRecordMapper outWorkRecordMapper;

    @Autowired
    OvertimeRecordMapper overtimeRecordMapper;

    @Autowired
    WorkRecordMapper workRecordMapper;

    @Autowired
    StatisticsResultMapper statisticsResultMapper;

  /* @Autowired
    EntityManagerUtil entityManagerUtil;*/

    @Override
    public UserDetails loadUserByUsername(String s) throws UsernameNotFoundException {
        User user = userMapper.loadUserByUsername(s);
        if (user == null) {
            //避免返回null，这里返回一个不含有任何值的User对象，在后期的密码比对过程中一样会验证失败
            return new User();
        }
        //查询用户的角色信息，并返回存入user中
        List<Role> roles = rolesMapper.getRolesByUid(user.getId());
        user.setRoles(roles);
        return user;
    }

    /**
     * @param user
     * @return 0表示成功
     * 1表示用户名重复
     * 2表示失败
     */
    public int reg(User user) {
        User loadUserByUsername = userMapper.loadUserByUsername(user.getUsername());
        if (loadUserByUsername != null) {
            return 1;
        }
        //插入用户,插入之前先对密码进行加密
        user.setPassword(passwordEncoder.encode(user.getPassword()));
        user.setEnabled(true);//用户可用
        long result = userMapper.reg(user);
        //配置用户的角色，默认都是普通用户
        String[] roles = new String[]{"2"};
        int i = rolesMapper.addRoles(roles, user.getId());
        boolean b = i == roles.length && result == 1;
        if (b) {
            return 0;
        } else {
            return 2;
        }
    }

    public int updateUserEmail(String email) {
        return userMapper.updateUserEmail(email, Util.getCurrentUser().getId());
    }

    public List<User> getUserByNickname(String nickname) {
        List<User> list = userMapper.getUserByNickname(nickname);
        return list;
    }

    public List<Role> getAllRole() {
        return userMapper.getAllRole();
    }

    public int updateUserEnabled(Boolean enabled, Long uid) {
        return userMapper.updateUserEnabled(enabled, uid);
    }

    public int deleteUserById(Long uid) {
        return userMapper.deleteUserById(uid);
    }

    public int updateUserRoles(Long[] rids, Long id) {
        int i = userMapper.deleteUserRolesByUid(id);
        return userMapper.setUserRoles(rids, id);
    }

    public User getUserById(Long id) {
        return userMapper.getUserById(id);
    }


    public List<PersonManage> getPersons(String name){
        return personManageMapper.getPersons(name);
    }
    public void deletePerson(String id){
        personManageMapper.deletePersons(id);
    }
    public void addOrEditPersons(PersonManage personManage){
        try {
            if(StringUtils.isEmpty(personManage.getId())){
                String maxid = personManageMapper.getMaxidPersons();
                String num = String.valueOf(Integer.parseInt(maxid)+1);
                int count  = 5-num.length();
                if(count !=0 ){
                    for(int i=0;i<count;i++){
                        num = '0'+num;
                    }
                }
                personManage.setId(num);
                personManageMapper.savePersons(personManage);
            }else {
                personManageMapper.updatePersons(personManage);
            }

        }catch (Exception e){
            e.printStackTrace();
        }
    }
    public PageObject<WorkRecord> getWorkRecords(String name,  int pageCount, int pageSize){
        //使用PageHelper封装分页信息
        Page<WorkRecord> page= PageHelper.startPage(pageCount,pageSize);
        List<WorkRecord> workRecordList =workRecordMapper.getWorkRecords(name);
        //PageObject为自定类，封装了查询到的分页内容
        return new PageObject((int)page.getTotal(), workRecordList, pageSize,pageCount);
    }
    public PageObject<OutWorkRecord>  getOutWorkRecords(String name,  int pageCount, int pageSize) throws ParseException {

        //使用PageHelper封装分页信息
        Page<OutWorkRecord> page= PageHelper.startPage(pageCount,pageSize);
        List<OutWorkRecord> outworkRecordList =outWorkRecordMapper.getOutWorkRecords(name);
        //PageObject为自定类，封装了查询到的分页内容
        return new PageObject((int)page.getTotal(), outworkRecordList, pageSize,pageCount);
    }
    public PageObject<Overtimerecord>  getOvertimeRecords(String name,  int pageCount, int pageSize){

        //使用PageHelper封装分页信息
        Page<Overtimerecord> page= PageHelper.startPage(pageCount,pageSize);
        List<Overtimerecord> overtimeRecordList =overtimeRecordMapper.getOvertimeRecords(name);
        //PageObject为自定类，封装了查询到的分页内容
        return new PageObject((int)page.getTotal(), overtimeRecordList, pageSize,pageCount);
    }

    public PageObject<StatisticsResult>  getStatisResult(String name,  int pageCount, int pageSize) throws ParseException {

        //使用PageHelper封装分页信息
        Page<StatisticsResult> page= PageHelper.startPage(pageCount,pageSize);
        List<StatisticsResult> statisticsResults =statisticsResultMapper.getStatisticsResultList(name);
        for(StatisticsResult statisticsResult:statisticsResults){
            String jb = setOvertime(statisticsResult.getTime(), statisticsResult.getName());
            statisticsResult.setOvertime(jb);
        }
        //PageObject为自定类，封装了查询到的分页内容
        return new PageObject((int)page.getTotal(), statisticsResults, pageSize,pageCount);
    }
    //处理加班
    public  String setOvertime( String time, String name) throws ParseException {
        String srt = "";
        SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy/MM/dd");
        List<Overtimerecord> overtimerecords =overtimeRecordMapper.getOvertimeRecords(name);
        for(Overtimerecord  ot : overtimerecords){
            String[] times1 = ot.getStarttime().split(" ");
            String[] times2 = ot.getEndtime().split(" ");
            Date startData = sdf1.parse(times1[0]);
            Date endData = sdf1.parse(times2[0]);
            Date time1 = sdf1.parse(time);
            if(time1.after(startData) && time1.before(endData)  || time1.equals(startData) || time1.equals(endData) ){
                srt = "1";
            }
            if((time1.equals(startData) && "下午".equals(times1[1])) || (time1.equals(endData) && "上午".equals(times2[1])) ){
                srt = "0.5";
            }
        }
        return srt;
    }

    @Transactional
    public void  doAttendanceStatistics() throws Exception{
        try {
            List<Map<String, Object>> mapList = statisticsResultMapper.getAllResultList();
            List<StatisticsResult> listResult = new ArrayList<>();
            statisticsResultMapper.delectStatisticResult();
            for(Map<String, Object> map:mapList){
                StatisticsResult sr = new StatisticsResult();
                if(map.get("name") != null){
                    sr.setName(map.get("name").toString());
                }
                if(map.get("time") != null){
                    sr.setTime(map.get("time").toString());
                }
                if(map.get("account") != null){
                    sr.setAccount(map.get("account").toString());
                }
                if(map.get("code") != null){
                    sr.setCode(map.get("code").toString());
                }
                if(map.get("apply") != null){
                    sr.setApplys(map.get("apply").toString());
                }
                if(map.get("shift") != null){
                    sr.setClasses(map.get("shift").toString());
                }
                if(map.get("firstWorkTime") != null){
                    sr.setFirst_work_time(map.get("firstWorkTime").toString());
                }
                if(map.get("lastWorkTime") != null){
                    sr.setLast_work_time(map.get("lastWorkTime").toString());
                }
                if(map.get("firstOutTime") != null){
                    sr.setOut_fist_time(map.get("firstOutTime").toString());
                }
                if(map.get("lastOutTime") != null){
                    sr.setOut_last_time(map.get("lastOutTime").toString());
                }
                if(map.get("outLocation") != null){
                    sr.setLocation(map.get("outLocation").toString());
                }
                sr.setNo_clock("0");
                sr.setPoints("0");
                if((map.get("apply") != null && map.get("apply").toString().contains("请假")) || (map.get("shift").toString().contains("休息")) || (map.get("shift").toString().contains("--"))){
                    sr.setNo_clock("0");
                    sr.setPoints("0");
                }else {

                   //打卡时间节点
                    SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
                    Date time1 = sdf.parse(map.get("shift").toString().substring(0, 5));
                    Date time2 = sdf.parse("12:00");
                    Date time3 = sdf.parse(map.get("shift").toString().substring(6));
                    Date time4 = sdf.parse("24:00");


                   //以下处理未打卡
                    if("未打卡".equals(map.get("firstWorkTime").toString()) && map.get("firstOutTime") == null){
                        sr.setNo_clock("1");
                    }

                    if("未打卡".equals(map.get("lastWorkTime").toString()) && map.get("firstOutTime") != null){
                        Date firstOutTime = sdf.parse(map.get("firstOutTime").toString());
                        if(firstOutTime.compareTo(time2)>0){
                            sr.setNo_clock("1");
                        }
                    }
                    if("未打卡".equals(map.get("lastWorkTime").toString()) && map.get("lastOutTime") != null){
                        Date lastOutTime = sdf.parse(map.get("lastOutTime").toString());
                        if(lastOutTime.compareTo(time2)<0){
                            sr.setNo_clock("1");
                        }
                    }

                    if("未打卡".equals(map.get("lastWorkTime").toString()) && map.get("lastOutTime") == null){
                        sr.setNo_clock("2");
                    }
                    //处理扣分
                        int num =0;
                        //上午
                        int num_am1 = 0;
                        int num_am2 = 0;
                        if( !"未打卡".equals(map.get("firstWorkTime").toString()) ){
                            Date firstWorkTime = sdf.parse(map.get("firstWorkTime").toString());
                            if(firstWorkTime.compareTo(time1)>0  && firstWorkTime.compareTo(time2)<=0){
                                num_am1 = minsBetween(firstWorkTime, time1 ) ;
                            }
                        }
                        if(map.get("firstOutTime") != null){
                            Date firstOutTime = sdf.parse(map.get("firstOutTime").toString());
                            if(firstOutTime.compareTo(time1)>0 && firstOutTime.compareTo(time2)<=0){
                                num_am2 = minsBetween(firstOutTime, time1 ) ;
                            }
                        }

                        num=num_am1>num_am2?num_am2:num_am1;//取最小值
                        //下午
                        int num_pm1 = 0;
                        int num_pm2 = 0;
                        if( !"未打卡".equals(map.get("lastWorkTime").toString()) ){
                            Date lasttWorkTime = sdf.parse(map.get("lastWorkTime").toString());
                            if(lasttWorkTime.compareTo(time3)<0 && lasttWorkTime.compareTo(time2)>0){
                                num_pm1 = minsBetween(lasttWorkTime, time1 ) + num;
                            }
                        }

                        if(map.get("lastOutTime") != null){
                            Date lastOutTime = sdf.parse(map.get("lastOutTime").toString());
                            if(lastOutTime.compareTo(time3)<0 && lastOutTime.compareTo(time2)>0){
                                num_pm2 = minsBetween(lastOutTime, time3 ) + num;
                            }
                        }
                        num=(num_pm1>num_pm2?num_pm2:num_pm1)  + num;//取最小值
                        sr.setPoints(Integer.toString(num));


                }
                //旷工处理
                if(StringUtils.isNotBlank(sr.getNo_clock()) && !"0".equals(sr.getNo_clock())){
                    sr.setAbsenteeism("1");
                }
                /*以下处理审批*/
             /*   List<String> applyList = new ArrayList<>();
                if(map.get("apply") != null  ){
                  String[] apply = map.get("apply").toString().split("、");
                  applyList = Arrays.asList(apply);
                }
                int count = 0;
                for(String apply :applyList){
                    String res = calculate(map.get("time").toString(), apply);
                    if(apply.contains("请假")){
                        sr.setVacate(res);
                    }
                    if(apply.contains("出差")){
                        sr.setEvection(res);
                    }
                    if(apply.contains("外出")){
                        sr.setGoout(res);
                    }
                    if(apply.contains("补卡")){
                        ++count;
                        sr.setReplenish(String.valueOf(count));
                    }

                }*/

                String str = UUID.randomUUID().toString().replace("-", "").toUpperCase();
                sr.setId(str);
                statisticsResultMapper.insertStatisticResult(sr);
            }
        }catch (Exception e){
            e.printStackTrace();
        }

    }
@Transactional
    public String calculate(String time, String apply) throws ParseException {
        List<String> stringList = Arrays.asList(apply.split(" "));
        String year = time.substring(0, 4);
        SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy/MM/dd");
        String starttime = "";
        String endtime = "";
        String startdate = "";
        String enddate = "";
        String result = "";

/*        if(apply.contains("请假") || apply.contains("出差" ) ){
            startdate = stringList.get(0);
            starttime = stringList.get(1);
            enddate = stringList.get(3);
            endtime = stringList.get(4);
            Date date1 = sdf1.parse(time);
            Date date2 = sdf1.parse(year + "/" + startdate);
            Date date3 = sdf1.parse(year + "/" + enddate);
            if((date1.after(date2) && date1.before(date3)) ||  date1.equals(date2) || date1.equals(date3)){
                result = "1";
            }
            if(  (date1.equals(date2) &&   starttime.contains("下午")) || (endtime.contains("上午") && date1.equals(date3) )){
                result = "0.5";
            }
        }*/
        /*if(apply.contains("外出")  ){
            startdate = stringList.get(0);
            starttime = stringList.get(1);
            enddate = stringList.get(3);
            endtime = stringList.get(4);
            Date date1 = sdf1.parse(time);
            Date date2 = sdf1.parse(year + "/" + startdate);
            Date date3 = sdf1.parse(year + "/" + enddate);
            if((date1.after(date2) && date1.before(date3)) ||  date1.equals(date2) || date1.equals(date3)){
                result = "1";
            }
            if(starttime.contains("下午")|| starttime.contains("上午")){
                if(  (date1.equals(date2) &&   starttime.contains("下午")) || (endtime.contains("上午") && date1.equals(date3) )){
                    result = "0.5";
                }
            }else {
                if(  (date1.equals(date2) && Integer.parseInt(starttime.substring(0,2))>=12 ) || ( date1.equals(date3) && Integer.parseInt(endtime.substring(0,2))<12 )){
                    result = "0.5";
                }
            }

        }*/
        return result;
    }

    /**
     *
     * @param dkTime  打卡时间
     * @param sbTime  上班时间
     * @return
     */
    public  int minsBetween(Date dkTime, Date sbTime) {
        Calendar cal = Calendar.getInstance();
        if (dkTime == null || sbTime == null) {
            return 0;
        }
        cal.setTime(dkTime);
        long time1 = cal.getTimeInMillis();
        cal.setTime(sbTime);
        long time2 = cal.getTimeInMillis();
        //算上当天
        int longTime = Math.abs(Integer.parseInt(String.valueOf((time1 - time2) / 60000L)) + 1);
        int num =0;
        if(longTime<=5 && longTime>0){
            num = 1;
        }else if(longTime<=15 && longTime>5){
            num = 3;
        }else if(longTime<=30 && longTime>15){
            num = 5;
        }else if(longTime<=60 && longTime>30){
            num = 10;
        }else {
            num = 20;
        }
        return num;
    }

    public List<Map<String, Object>>  doAttendanceStatisticsByMonth(String name) throws Exception{
        List<Map<String, Object>> mapList = new ArrayList<>();
        try {
           mapList = statisticsResultMapper.getAllResultListByMonth(name);

        }catch (Exception e){
            e.printStackTrace();
        }
        return mapList;
    }


    public void  exportAttendance(HttpServletResponse response) throws Exception{
        List<Map<String, Object>> mapList = new ArrayList<>();
        try {
            String name = "";
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
            Date d = new Date();
            String fileName = sdf.format(d);
            String outPath = "D:\\kqSystem\\data\\upload\\template\\本月考勤统计("+fileName+").xlsx" ;
            String[] heads = {"日期", "名称", "班次", "上班最早", "下班最晚", "外出最早","外出最晚", "迟到早退扣分","外出打卡地点", "考勤申请","迟到早退扣分","未打卡次数", "出差","加班","外出","补打卡","旷工","请假","校对状态"};
            List<Map<String,Object>> statisticsResults =statisticsResultMapper.getStatisticsResultListToExport(name);
            for(Map<String,Object> map: statisticsResults){
                String jbdata=setOvertime(map.get("日期").toString(), map.get("名称").toString());
                map.put("overtime", jbdata);
            }
            File file = new File(outPath);
            if(!file.exists()){
                file.createNewFile();
            }
            ExcelUtil.exportExcel(heads,statisticsResults,outPath);
            //下载
            ExcelUtil.download(outPath, response );

        }catch (Exception e){
            e.printStackTrace();
        }
    }
    public void  exportAttendanceByMonth(HttpServletResponse response) throws Exception{
        List<Map<String, Object>> mapList = new ArrayList<>();
        try {
            String name = "";
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
            Date d = new Date();
            String fileName = sdf.format(d);
            String outPath = "D:\\kqSystem\\data\\upload\\template\\本月考勤汇总统计("+fileName+").xlsx" ;
            String[] heads = {"考勤号", "名称", "账号", "本月未打卡次数", "本月扣分总和"};
            List<Map<String,Object>> statisticsResults =statisticsResultMapper.getAllResultListByMonthToExport(name);
            File file = new File(outPath);
            if(!file.exists()){
                file.createNewFile();
            }
            ExcelUtil.exportExcel(heads,statisticsResults,outPath);
            //下载
            ExcelUtil.download(outPath, response );
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    public void  AttendanceForApply(HttpServletResponse response) throws Exception{
        try {
            List<StatisticsResult> statisticsResults =statisticsResultMapper.getStatisticsResultList("");
            File fileA = new File("D:\\kqSystem\\data\\upload\\template\\考勤申请记录模板.xls");
            Date d = new Date();
            SimpleDateFormat sdfDefine = new SimpleDateFormat("HH:mm");
            Date definetime = sdfDefine.parse("12:00");
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
            String fileName = sdf.format(d);
            String outPath = "D:\\kqSystem\\data\\upload\\template\\" + "考勤申请"+fileName + ".xls";
            if(!fileA.exists()){
                //如果文件存在就删除
                throw new   RuntimeException("没有模板文件！");
            }
            fileA.getName();
            //获得文件的后缀名
            FileInputStream inputStream = new FileInputStream(fileA);
            //根据后缀名判断EE
            Workbook wb = ExcelUtil.getWorkbook(fileA);
            CellStyle style = wb.createCellStyle();
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            CellStyle setBorder = wb.createCellStyle();
            setBorder.setBorderBottom(CellStyle.BORDER_THIN); //下边框
            setBorder.setBorderLeft(CellStyle.BORDER_THIN);//左边框
            setBorder.setBorderTop(CellStyle.BORDER_THIN);//上边框
            setBorder.setBorderRight(CellStyle.BORDER_THIN);//右边框
            // HSSFWorkbook wb = new HSSFWorkbook(inputStream);
            Sheet sheet = wb.getSheetAt(0);
            inputStream.close();
            sheet.setForceFormulaRecalculation(true);//强制执行excel中函数
           // int rows = sheet.getLastRowNum()+1;
            int rowNum = 1;
            int index =0;
            int num =1;
            String Account = "";
            Row newRow = null;
            String qjAll = "0";
            String jbAll = "0";
            String wcAll = "0";
            String ccAll = "0";
            String bdkAll = "0";
            String cdAll = "0";
            String kgAll = "0";
            for(StatisticsResult sr :statisticsResults){
                Cell cell = null;
                if(!Account.equals(sr.getAccount()) ){
                    if(rowNum != 1){
                        newRow = sheet.getRow(rowNum);
                        //请假
                        cell = newRow.getCell(218);
                        cell.setCellValue(qjAll);
                        //加班
                        cell = newRow.getCell(219);
                        cell.setCellValue(jbAll);
                        //外出
                        cell = newRow.getCell(220);
                        cell.setCellValue(wcAll);
                        //出差
                        cell = newRow.getCell(221);
                        cell.setCellValue(ccAll);
                        //补打卡
                        cell = newRow.getCell(222);
                        cell.setCellValue(bdkAll);
                        //迟到
                        cell = newRow.getCell(223);
                        cell.setCellValue(cdAll);
                        //旷工
                        cell = newRow.getCell(224);
                        cell.setCellValue(kgAll);

                        qjAll = "0";
                        jbAll = "0";
                        wcAll = "0";
                        ccAll = "0";
                        bdkAll = "0";
                        cdAll = "0";
                        kgAll = "0";
                        num=1;
                    }
                    Account = sr.getAccount();
                    index =0;
                    rowNum++ ;
                    newRow = sheet.getRow(rowNum);
                }

                //是否从月初第一天开始考勤
                if(Integer.parseInt(sr.getTime().substring(sr.getTime().length()-2)) != num){
                    num = Integer.parseInt(sr.getTime().substring(sr.getTime().length()-2));
                    index = (int)(num-1)* 7;
                }

                //名称
                cell = newRow.getCell(0);
                cell.setCellValue(sr.getName());
                //请假
                cell = newRow.getCell(index +1);
                cell.setCellValue(sr.getVacate());
                String value = StringUtils.isBlank(sr.getVacate()) ? "0": sr.getVacate();
                qjAll=new BigDecimal(qjAll).add(new BigDecimal(value)).toString();
                //加班
                cell = newRow.getCell(index +2);
                String jb = setOvertime(sr.getTime(), sr.getName());
                cell.setCellValue(jb);
                value = StringUtils.isBlank(jb) ? "0": jb;
                jbAll=new BigDecimal(jbAll).add(new BigDecimal(value)).toString();
                //外出
                cell = newRow.getCell(index +3);
                cell.setCellValue(sr.getGoout());
                value = StringUtils.isBlank(sr.getGoout()) ? "0": sr.getGoout();
                wcAll=new BigDecimal(wcAll).add(new BigDecimal(value)).toString();
                //出差
                cell = newRow.getCell(index +4);
                cell.setCellValue(sr.getEvection());
                value = StringUtils.isBlank(sr.getEvection()) ? "0": sr.getEvection();
                ccAll=new BigDecimal(ccAll).add(new BigDecimal(value)).toString();
                //补打卡
                cell = newRow.getCell(index +5);
                cell.setCellValue(sr.getReplenish());
                value = StringUtils.isBlank(sr.getReplenish()) ? "0": sr.getReplenish();
                bdkAll=new BigDecimal(bdkAll).add(new BigDecimal(value)).toString();
                //迟到
                cell = newRow.getCell(index +6);
                if(!"0".equals(sr.getPoints())){
                    cell.setCellValue(sr.getPoints());
                }
                cdAll=new BigDecimal(cdAll).add(new BigDecimal(sr.getPoints())).toString();
                //旷工
                cell = newRow.getCell(index +7);
                cell.setCellValue(sr.getAbsenteeism());
                value = StringUtils.isBlank(sr.getAbsenteeism()) ? "0": sr.getAbsenteeism();
                kgAll=new BigDecimal(kgAll).add(new BigDecimal(value)).toString();
                index = index + 7 ;
                num++;
                //合计最后一行
                if( sr.equals(statisticsResults.get(statisticsResults.size()-1))){
                    if(rowNum != 1){
                        //请假
                        cell = newRow.getCell(218);
                        cell.setCellValue(qjAll);
                        //加班
                        cell = newRow.getCell(219);
                        cell.setCellValue(jbAll);
                        //外出
                        cell = newRow.getCell(220);
                        cell.setCellValue(wcAll);
                        //出差
                        cell = newRow.getCell(221);
                        cell.setCellValue(ccAll);
                        //补打卡
                        cell = newRow.getCell(222);
                        cell.setCellValue(bdkAll);
                        //迟到
                        cell = newRow.getCell(223);
                        cell.setCellValue(cdAll);
                        //旷工
                        cell = newRow.getCell(224);
                        cell.setCellValue(kgAll);

                    }
                }
            }


            FileOutputStream excelFileOutPutStream = new FileOutputStream(outPath);
            // 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
            wb.write(excelFileOutPutStream);
            // 执行 flush 操作， 将缓存区内的信息更新到文件上
            excelFileOutPutStream.flush();
            // 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
            excelFileOutPutStream.close();

            //下载
            ExcelUtil.download(outPath, response );
            File file = new File(outPath);
            file.delete();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    public void  exportAttendanceByRule(HttpServletResponse response) throws Exception{
        try {
            List<StatisticsResult> statisticsResults =statisticsResultMapper.getStatisticsResultList("");
            File fileA = new File("D:\\kqSystem\\data\\upload\\template\\考勤导入模板.xls");
            Date d = new Date();
            SimpleDateFormat sdfDefine = new SimpleDateFormat("HH:mm");
            Date definetime = sdfDefine.parse("12:00");
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
            String fileName = sdf.format(d);
            String outPath = "D:\\kqSystem\\data\\upload\\template\\" + "考勤导入"+fileName + ".xls";
            if(!fileA.exists()){
                //如果文件存在就删除
                throw new   RuntimeException("没有模板文件！");
            }
            fileA.getName();
            //获得文件的后缀名
            FileInputStream inputStream = new FileInputStream(fileA);
            //根据后缀名判断EE
            Workbook wb = ExcelUtil.getWorkbook(fileA);
            CellStyle style = wb.createCellStyle();
            style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            // HSSFWorkbook wb = new HSSFWorkbook(inputStream);
            Sheet sheet = wb.getSheetAt(0);
            inputStream.close();
            sheet.setForceFormulaRecalculation(true);//强制执行excel中函数
           // int rows = sheet.getLastRowNum()+1;
            int rowNum = 1;
            for(StatisticsResult sr :statisticsResults){
                    Row newRow = sheet.createRow(rowNum);
                    Cell cell = null;

                    if( (sr.getClasses().contains("休息")) || (sr.getClasses().contains("--")) || (("未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() == null) && "未打卡".equals(sr.getLast_work_time())  && sr.getOut_last_time() == null )) {
                        continue;
                        // }else if( sr.getApplys() != null && sr.getApplys().contains("请假")){
                    }else {
                        cell = newRow.createCell(0, Cell.CELL_TYPE_STRING);
                        cell.setCellValue(sr.getCode());
                        cell = newRow.createCell(1, Cell.CELL_TYPE_STRING);
                        cell.setCellValue(sr.getName());
                        cell = newRow.createCell(2, Cell.CELL_TYPE_STRING);
                        cell.setCellValue(sr.getTime());
                        if("10:00-20:00".equals(sr.getClasses()) || "10:00-19:30".equals(sr.getClasses())){
                            //上午打卡(新疆班次)
                            if("未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() == null){
                                cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                cell.setCellValue("");
                            }else
                            if(!"未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() == null){
                                Date firstWorkTime = sdfDefine.parse(sr.getFirst_work_time());
                                Calendar c = Calendar.getInstance();
                                c.setTime(firstWorkTime);
                                c.add(Calendar.HOUR_OF_DAY, -1);
                                firstWorkTime = c.getTime();
                                if(definetime.compareTo(firstWorkTime)>=0){
                                    cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                    cell.setCellValue(sdfDefine.format(firstWorkTime));
                                }
                            }else
                            if("未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() != null){
                                Date lastWorkTime = sdfDefine.parse(sr.getOut_fist_time());
                                Calendar c = Calendar.getInstance();
                                c.setTime(lastWorkTime);
                                c.add(Calendar.HOUR_OF_DAY, -1);
                                lastWorkTime = c.getTime();
                                if(definetime.compareTo(lastWorkTime)>=0){
                                    cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                    cell.setCellValue(sdfDefine.format(lastWorkTime));
                                }
                            }else
                            if(!"未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() != null){
                                Date firstWorkTime = sdfDefine.parse(sr.getFirst_work_time());
                                Date firstOutTime = sdfDefine.parse(sr.getOut_fist_time());
                                String time = firstWorkTime.compareTo(firstOutTime) <=0 ? sr.getFirst_work_time(): sr.getOut_fist_time();
                                Date timeDate = sdfDefine.parse(time);
                                Calendar c = Calendar.getInstance();
                                c.setTime(timeDate);
                                c.add(Calendar.HOUR_OF_DAY, -1);
                                timeDate = c.getTime();
                                if(definetime.compareTo(timeDate)>=0){
                                    cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                    cell.setCellValue(sdfDefine.format(timeDate));
                                }
                            }
                        }else {
                            //上午打卡
                            if("未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() == null){
                                cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                cell.setCellValue("");
                            }else
                            if(!"未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() == null){
                                Date firstWorkTime = sdfDefine.parse(sr.getFirst_work_time());
                                if(definetime.compareTo(firstWorkTime)>=0){
                                    cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                    cell.setCellValue(sr.getFirst_work_time());
                                }
                            }else
                            if("未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() != null){
                                Date lastWorkTime = sdfDefine.parse(sr.getOut_fist_time());
                                if(definetime.compareTo(lastWorkTime)>=0){
                                    cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                    cell.setCellValue(sr.getOut_fist_time());
                                }
                            }else
                            if(!"未打卡".equals(sr.getFirst_work_time())  && sr.getOut_fist_time() != null){
                                Date firstWorkTime = sdfDefine.parse(sr.getFirst_work_time());
                                Date firstOutTime = sdfDefine.parse(sr.getOut_fist_time());
                                String time = firstWorkTime.compareTo(firstOutTime) <=0 ? sr.getFirst_work_time(): sr.getOut_fist_time();
                                Date timeDate = sdfDefine.parse(time);
                                if(definetime.compareTo(timeDate)>=0){
                                    cell = newRow.createCell(3, Cell.CELL_TYPE_STRING);
                                    cell.setCellValue(time);
                                }
                            }
                        }




                        //下午打卡
                        if("未打卡".equals(sr.getLast_work_time())  && sr.getOut_last_time() == null){
                            cell = newRow.createCell(4, Cell.CELL_TYPE_STRING);
                            cell.setCellValue("");
                        }else
                        if(!"未打卡".equals(sr.getLast_work_time())  && sr.getOut_last_time() == null){
                            Date lastWorkTime = sdfDefine.parse(sr.getLast_work_time());
                            if(definetime.compareTo(lastWorkTime)<=0){
                                cell = newRow.createCell(4, Cell.CELL_TYPE_STRING);
                                cell.setCellValue(sr.getLast_work_time());
                            }
                        }else
                        if("未打卡".equals(sr.getLast_work_time())  && sr.getOut_last_time() != null){
                            Date lastOutTime = sdfDefine.parse(sr.getOut_last_time());
                            if(definetime.compareTo(lastOutTime)<=0){
                                cell = newRow.createCell(4, Cell.CELL_TYPE_STRING);
                                cell.setCellValue(sr.getOut_last_time());
                            }
                        }else
                        if(!"未打卡".equals(sr.getLast_work_time())  && sr.getOut_last_time() != null){
                            Date lastWorkTime = sdfDefine.parse(sr.getLast_work_time());
                            Date lastOutTime = sdfDefine.parse(sr.getOut_last_time());
                            String time = lastWorkTime.compareTo(lastOutTime) >=0 ? sr.getLast_work_time(): sr.getOut_last_time();
                            Date timeDate = sdfDefine.parse(time);
                            if(definetime.compareTo(timeDate)<=0){
                                cell = newRow.createCell(4, Cell.CELL_TYPE_STRING);
                                cell.setCellValue(time);
                            }
                        }
                        ++ rowNum ;
                    }

                }
            FileOutputStream excelFileOutPutStream = new FileOutputStream(outPath);
            // 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
            wb.write(excelFileOutPutStream);
            // 执行 flush 操作， 将缓存区内的信息更新到文件上
            excelFileOutPutStream.flush();
            // 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
            excelFileOutPutStream.close();
            //下载
            ExcelUtil.download(outPath, response );
            File file = new File(outPath);
            file.delete();
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
