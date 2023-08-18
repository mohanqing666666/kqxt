package org.sang.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.sang.bean.StatisticsResult;
import org.sang.bean.WorkRecord;

import java.util.List;

@Mapper
public interface WorkRecordMapper {
    List<WorkRecord> getWorkRecords(@Param("name") String name);
    int  delectWorkRecord();
    int insertWorkRecord(WorkRecord workRecord);
}
