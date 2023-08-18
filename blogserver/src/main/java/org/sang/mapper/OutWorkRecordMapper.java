package org.sang.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.sang.bean.OutWorkRecord;
import org.sang.bean.WorkRecord;

import java.util.List;

@Mapper
public interface OutWorkRecordMapper {
    List<OutWorkRecord> getOutWorkRecords(@Param("name") String name);
    int  delectOutWorkRecord();
    int insertOutWorkRecord(OutWorkRecord outWorkRecord);
}
