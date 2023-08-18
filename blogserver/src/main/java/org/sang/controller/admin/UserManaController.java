package org.sang.controller.admin;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.sang.bean.*;
import org.sang.mapper.OutWorkRecordMapper;
import org.sang.mapper.OvertimeRecordMapper;
import org.sang.mapper.WorkRecordMapper;
import org.sang.service.UserService;
import org.sang.utils.ExcelUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by sang on 2017/12/24.
 */
@RestController
@RequestMapping("/admin")
public class UserManaController {
    @Autowired
    UserService userService;

    @Autowired
    WorkRecordMapper workRecordMapper;

    @Autowired
    OutWorkRecordMapper outWorkRecordMapper;

    @Autowired
    OvertimeRecordMapper overtimeRecordMapper;

    @RequestMapping(value = "/user", method = RequestMethod.GET)
    public List<User> getUserByNickname(String nickname) {
        return userService.getUserByNickname(nickname);
    }

    @RequestMapping(value = "/user/{id}", method = RequestMethod.GET)
    public User getUserById(@PathVariable Long id) {
        return userService.getUserById(id);
    }

    @RequestMapping(value = "/roles", method = RequestMethod.GET)
    public List<Role> getAllRole() {
        return userService.getAllRole();
    }

    @RequestMapping(value = "/user/enabled", method = RequestMethod.PUT)
    public RespBean updateUserEnabled(Boolean enabled, Long uid) {
        if (userService.updateUserEnabled(enabled, uid) == 1) {
            return new RespBean("success", "更新成功!");
        } else {
            return new RespBean("error", "更新失败!");
        }
    }

    @RequestMapping(value = "/user/{uid}", method = RequestMethod.DELETE)
    public RespBean deleteUserById(@PathVariable Long uid) {
        if (userService.deleteUserById(uid) == 1) {
            return new RespBean("success", "删除成功!");
        } else {
            return new RespBean("error", "删除失败!");
        }
    }

    @RequestMapping(value = "/user/role", method = RequestMethod.PUT)
    public RespBean updateUserRoles(Long[] rids, Long id) {
        if (userService.updateUserRoles(rids, id) == rids.length) {
            return new RespBean("success", "更新成功!");
        } else {
            return new RespBean("error", "更新失败!");
        }
    }

    /**
     * 获取人员列表
     * @param name
     * @return
     */
    @RequestMapping(value = "/getPersonList" , method = RequestMethod.GET)
    public List<PersonManage> getPersonList(String name) {
        return userService.getPersons(name);
    }
    /**
     * 新增、编辑人员信息
     * @return
     */
    @RequestMapping(value = "/addOrEditPerson" , method = RequestMethod.POST)
    public JsonResult addOrEditPerson(PersonManage person) {
        JsonResult jsonResult = new JsonResult();
        try {
            userService.addOrEditPersons(person);
            jsonResult.setOk("OK");
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setMsg(e.getMessage());
            jsonResult.setOk("err");
        }
        return jsonResult;
    }

    /**
     * 删除人员信息
     * @return
     */
    @RequestMapping(value = "/deletePerson" , method = RequestMethod.POST)
    public JsonResult addOrEditPerson(String id) {
        JsonResult jsonResult = new JsonResult();
        try {
            userService.deletePerson(id);
            jsonResult.setOk("OK");
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setMsg(e.getMessage());
            jsonResult.setOk("err");
        }
        return jsonResult;
    }

    /**
     * 获取正常打卡记录
     * @param name
     * @return
     */
    @RequestMapping(value = "/getWorkRecordList" , method = RequestMethod.GET)
    public JsonResult  getWorkRecordList(String name, int pageCount, int pageSize) {
        PageObject<WorkRecord> pageObject =  userService.getWorkRecords(name,   pageCount,  pageSize);
        return new JsonResult(pageObject);
    }

    /**
     * 获取外出打卡记录
     * @param name
     * @return
     */
    @RequestMapping(value = "/getOutWorkRecordList" , method = RequestMethod.GET)
    public JsonResult getOutWorkRecordList(String name, int pageCount, int pageSize) throws ParseException {

        PageObject<OutWorkRecord> pageObject =  userService.getOutWorkRecords(name,   pageCount,  pageSize);
        return new JsonResult(pageObject);
    }
    /**
     * 获取加班申请记录
     * @param name
     * @return
     */
    @RequestMapping(value = "/getOvertimeRecordList" , method = RequestMethod.GET)
    public JsonResult getOvertimeRecordList(String name, int pageCount, int pageSize) {

        PageObject<Overtimerecord> pageObject =  userService.getOvertimeRecords(name,   pageCount,  pageSize);
        return new JsonResult(pageObject);
    }

    /**
     * 获取考勤计算结果
     * @param name
     * @return
     */
    @RequestMapping(value = "/getStatisticsResultList" , method = RequestMethod.GET)
    public JsonResult getStatisticsResultList(String name, int pageCount, int pageSize) {
        JsonResult jsonResult = new JsonResult();
        try{
            PageObject<StatisticsResult> pageObject =  userService.getStatisResult(name,   pageCount,  pageSize);
            jsonResult = new JsonResult(pageObject);
        }catch (Exception e){
            jsonResult.setMsg(e.getMessage());
            jsonResult.setOk("err");
            e.printStackTrace();
        }

        return jsonResult;
    }

    /**
     * 考勤计算计算（12）
     * @return
     */
    @RequestMapping(value = "/doAttendanceStatistics" , method = RequestMethod.POST)
    public JsonResult doAttendanceStatistics() {
        JsonResult jsonResult = new JsonResult();
        try {
            userService.doAttendanceStatistics();
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setMsg(e.getMessage());
        }
        jsonResult.setOk("OK");
        return jsonResult;
    }

    /**
     * 考勤计算计算
     * @param name
     * @return
     */
    @RequestMapping(value = "/doAttendanceStatisticsByMonth" , method = RequestMethod.GET)
    public JsonResult doAttendanceStatisticsByMonth(String name) {
        JsonResult jsonResult = new JsonResult();
        try {
            jsonResult.setData(userService.doAttendanceStatisticsByMonth(name));
            jsonResult.setOk("OK");
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setOk("err");
            jsonResult.setMsg(e.getMessage());
        }

        return jsonResult;
    }

    /**
     * 考勤导出
     * @return
     */
    @RequestMapping(value = "/exportAttendance" , method = RequestMethod.GET)
    public void exportAttendance(HttpServletResponse response) {
        try {
            userService.exportAttendance(response);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    /**
     * 按月考勤导出
     * @return
     */
    @RequestMapping(value = "/exportAttendanceByMonth" , method = RequestMethod.GET)
    public void exportAttendanceByMonth(HttpServletResponse response) {
        try {
            userService.exportAttendanceByMonth(response);
        }catch (Exception e){
            e.printStackTrace();
        }

    }

    /**
     * 入库导出
     * @return
     */
    @RequestMapping(value = "/exportAttendanceByRule" , method = RequestMethod.GET)
    public void exportAttendanceByRule(HttpServletResponse response) {
        try {
            userService.exportAttendanceByRule(response);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    /**
     * 考勤申请情况导出
     * @return
     */
    @RequestMapping(value = "/exportAttendanceForApply" , method = RequestMethod.GET)
    public void exportAttendanceForApply(HttpServletResponse response) {
        try {
            userService.AttendanceForApply(response);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
    /**
     * 文件上传（上下班打卡）
     * @return
     */
    @PostMapping("/upload")
    @Transactional
    public JsonResult upload(@RequestParam MultipartFile file) {
        JsonResult jsonResult = new JsonResult();
        try {
            if(file.isEmpty()){
                jsonResult.setMsg("文件为空！");
                jsonResult.setOk("err");
                throw new RuntimeException("文件为空");
            }
            //根据路径获取这个操作excel的实例
            //HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
            //获取文件的原名称 getOriginalFilename
            String OriginalFilename = file.getOriginalFilename();
            //获取时间戳和文件的扩展名，拼接成一个全新的文件名； 用时间戳来命名是为了避免文件名冲突
            String fileName = System.currentTimeMillis()+"."+OriginalFilename.substring(OriginalFilename.lastIndexOf(".")+1);
            //定义文件存放路径
            String filePath = "D:\\kqSystem\\data\\upload\\";
            //新建一个目录（文件夹）
            File dest = new File(filePath+fileName);
            //判断filePath目录是否存在，如不存在，就新建一个
            if (!dest.getParentFile().canExecute()){
                dest.getParentFile().mkdirs(); //新建一个目录
            }
            file.transferTo(dest);
            Workbook wb = ExcelUtil.getWorkbook(dest);
            //根据页面index 获取sheet页
            Sheet sheet = wb.getSheetAt(0);
            Row row = null;
            //循环sesheet页中数据从第二行开始，第一行是标题
            List<WorkRecord> list = new ArrayList<>();
            workRecordMapper.delectWorkRecord();
            for (int i = 4; i < sheet.getPhysicalNumberOfRows(); i++) {
                //获取每一行数据
                row = sheet.getRow(i);
                WorkRecord wr = new WorkRecord();
                if(row.getCell(0) != null){
                    wr.setTime(row.getCell(0).toString());
                }
                if(row.getCell(1) != null){
                    wr.setName(row.getCell(1).toString());
                }
                if(row.getCell(2) != null){
                    wr.setAccount(row.getCell(2).toString());
                }
                if(row.getCell(3) != null){
                    wr.setRule(row.getCell(3).toString());
                }
                if(row.getCell(4) != null){
                    wr.setShift(row.getCell(4).toString());
                }
                if(row.getCell(5) != null){
                    wr.setFirst_work_time(row.getCell(5).toString());
                }
                if(row.getCell(6) != null){
                    wr.setLast_work_time(row.getCell(6).toString());
                }
                if(row.getCell(7) != null){
                    wr.setApply(row.getCell(7).toString());
                }
                if(row.getCell(8) != null){
                    wr.setCorrect_status(row.getCell(8).toString());
                }
                if(row.getCell(9) != null){
                    wr.setOut_count(row.getCell(9).toString());
                }
                if(row.getCell(10) != null){
                    wr.setOut_long(row.getCell(10).toString());
                }
                if(row.getCell(11) != null){
                    wr.setEvection_day(row.getCell(11).toString());
                }
                if(row.getCell(12) != null){
                    wr.setClock_in_time(row.getCell(12).toString());
                }
                if(row.getCell(13) != null){
                    wr.setClock_in_status(row.getCell(13).toString());
                }
                if(row.getCell(14) != null){
                    wr.setClock_in_time1(row.getCell(14).toString());
                }
                if(row.getCell(15) != null){
                    wr.setClock_in_status1(row.getCell(15).toString());
                }
                workRecordMapper.insertWorkRecord(wr);
            }
            dest.delete();//删除文件
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setMsg(e.getMessage());
            jsonResult.setOk("err");
            throw new RuntimeException(e.getMessage());
        }
        return jsonResult;

    }


    /**
     * 文件上传（外出打卡）
     * @return
     */
    @PostMapping("/upload2")
    @Transactional
    public JsonResult upload2(@RequestParam MultipartFile file) {
        JsonResult jsonResult = new JsonResult();
        try {
            if(file.isEmpty()){
                jsonResult.setMsg("文件为空！");
                jsonResult.setOk("err");
                throw new RuntimeException("文件为空");
            }
            //根据路径获取这个操作excel的实例
            //HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
            //获取文件的原名称 getOriginalFilename
            String OriginalFilename = file.getOriginalFilename();
            //获取时间戳和文件的扩展名，拼接成一个全新的文件名； 用时间戳来命名是为了避免文件名冲突
            String fileName = System.currentTimeMillis()+"."+OriginalFilename.substring(OriginalFilename.lastIndexOf(".")+1);
            //定义文件存放路径
            String filePath = "D:\\kqSystem\\data\\upload\\";
            //新建一个目录（文件夹）
            File dest = new File(filePath+fileName);
            //判断filePath目录是否存在，如不存在，就新建一个
            if (!dest.getParentFile().canExecute()){
                dest.getParentFile().mkdirs(); //新建一个目录
            }
            file.transferTo(dest);
            Workbook wb = ExcelUtil.getWorkbook(dest);
            //根据页面index 获取sheet页
            Sheet sheet = wb.getSheetAt(0);
            Row row = null;
            //循环sesheet页中数据从第二行开始，第一行是标题
            List<WorkRecord> list = new ArrayList<>();
            outWorkRecordMapper.delectOutWorkRecord();
            for (int i = 3; i < sheet.getPhysicalNumberOfRows(); i++) {
                //获取每一行数据
                row = sheet.getRow(i);
                OutWorkRecord owr = new OutWorkRecord();
                if(row.getCell(0) != null){
                    owr.setTime(row.getCell(0).toString());
                }
                if(row.getCell(1) != null){
                    owr.setName(row.getCell(1).toString());
                }
                if(row.getCell(2) != null){
                    owr.setAccount(row.getCell(2).toString());
                }
                if(row.getCell(3) != null){
                    owr.setDepartment(row.getCell(3).toString());
                }
                if(row.getCell(4) != null){
                    owr.setFirst_work_time(row.getCell(4).toString());
                }
                if(row.getCell(5) != null){
                    owr.setLast_work_time(row.getCell(5).toString());
                }
                if(row.getCell(6) != null){
                    owr.setClock_in_count(row.getCell(6).toString());
                }
                if(row.getCell(7) != null){
                    owr.setClock_in_time(row.getCell(7).toString());
                }
                if(row.getCell(8) != null){
                    owr.setClock_in_location(row.getCell(8).toString());
                }
                if(row.getCell(9) != null){
                    owr.setBz(row.getCell(9).toString());
                }
                if(row.getCell(10) != null){
                    owr.setClock_in_time1(row.getCell(10).toString());
                }
                if(row.getCell(11) != null){
                    owr.setClock_in_location1(row.getCell(11).toString());
                }
                if(row.getCell(12) != null){
                    owr.setBz1(row.getCell(12).toString());
                }
                if(row.getCell(13) != null){
                    owr.setClock_in_time2(row.getCell(13).toString());
                }
                if(row.getCell(14) != null){
                    owr.setClock_in_location2(row.getCell(14).toString());
                }
                if(row.getCell(15) != null){
                    owr.setBz2(row.getCell(15).toString());
                }
                outWorkRecordMapper.insertOutWorkRecord(owr);
            }
            dest.delete();//删除文件
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setMsg(e.getMessage());
            jsonResult.setOk("err");
            throw new RuntimeException(e.getMessage());
        }
        return jsonResult;

    }

    /**
     * 文件上传(加班申请)
     * @return
     */
    @PostMapping("/upload3")
    @Transactional
    public JsonResult upload3(@RequestParam MultipartFile file) {
        JsonResult jsonResult = new JsonResult();
        try {
            if(file.isEmpty()){
                jsonResult.setMsg("文件为空！");
                jsonResult.setOk("err");
                throw new RuntimeException("文件为空");
            }
            //根据路径获取这个操作excel的实例
            //HSSFWorkbook wb = new HSSFWorkbook(file.getInputStream());
            //获取文件的原名称 getOriginalFilename
            String OriginalFilename = file.getOriginalFilename();
            //获取时间戳和文件的扩展名，拼接成一个全新的文件名； 用时间戳来命名是为了避免文件名冲突
            String fileName = System.currentTimeMillis()+"."+OriginalFilename.substring(OriginalFilename.lastIndexOf(".")+1);
            //定义文件存放路径
            String filePath = "D:\\kqSystem\\data\\upload\\";
            //新建一个目录（文件夹）
            File dest = new File(filePath+fileName);
            //判断filePath目录是否存在，如不存在，就新建一个
            if (!dest.getParentFile().canExecute()){
                dest.getParentFile().mkdirs(); //新建一个目录
            }
            file.transferTo(dest);
            Workbook wb = ExcelUtil.getWorkbook(dest);
            //根据页面index 获取sheet页
            Sheet sheet = wb.getSheetAt(0);
            Row row = null;
            //循环sesheet页中数据从第二行开始，第一行是标题
            List<WorkRecord> list = new ArrayList<>();
            overtimeRecordMapper.delectOvertimeRecord();
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                //获取每一行数据
                row = sheet.getRow(i);
                if(StringUtils.isNotBlank(row.getCell(1).toString())){
                    Overtimerecord owr = new Overtimerecord();
                    owr.setType("加班");
                    if(row.getCell(0) != null){
                        owr.setCode(row.getCell(0).toString());
                    }
                    if(row.getCell(1) != null){
                        owr.setName(row.getCell(1).toString());
                    }
                    if(row.getCell(11) != null){
                        owr.setRecord(row.getCell(11).toString());
                    }
                    if(row.getCell(4) != null){
                        owr.setStarttime(row.getCell(4).toString()+ " " + row.getCell(5).toString());
                    }
                    if(row.getCell(6) != null){
                        owr.setEndtime(row.getCell(6).toString()+ " " + row.getCell(7).toString());
                    }
                    if(row.getCell(10) != null){
                        owr.setResultstatus(row.getCell(10).toString());
                    }
                    BigDecimal bigDecimal2 = new BigDecimal(row.getCell(13).toString());
                    if(row.getCell(13) != null && !bigDecimal2.equals(BigDecimal.ZERO)  ){
                        owr.setLongtime(row.getCell(13).toString());
                    }
                    BigDecimal bigDecimal1 = new BigDecimal(row.getCell(12).toString());
                    if(bigDecimal1 !=null && !bigDecimal1.equals(BigDecimal.ZERO) ){
                        bigDecimal2 = new BigDecimal(row.getCell(13).toString());
                        String time = bigDecimal1.add(bigDecimal2).toString();
                        owr.setLongtime( time);
                    }
                    overtimeRecordMapper.insertOvertimeRecord(owr);
                }
            }
            dest.delete();//删除文件
        }catch (Exception e){
            e.printStackTrace();
            jsonResult.setMsg(e.getMessage());
            jsonResult.setOk("err");
            throw new RuntimeException(e.getMessage());
        }
        return jsonResult;

    }

}
