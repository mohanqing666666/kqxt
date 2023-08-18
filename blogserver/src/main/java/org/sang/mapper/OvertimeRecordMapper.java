package org.sang.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.sang.bean.OutWorkRecord;
import org.sang.bean.Overtimerecord;

import java.util.List;

@Mapper
public interface OvertimeRecordMapper {
    List<Overtimerecord> getOvertimeRecords(@Param("name") String name);
    int  delectOvertimeRecord();
    int insertOvertimeRecord(Overtimerecord overtimerecord);
    List<Overtimerecord> getOvertimeRecordsByAccount(@Param("account") String account);
}
