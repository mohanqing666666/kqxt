package org.sang.mapper;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;
import org.sang.bean.PersonManage;

import java.util.List;

@Mapper
public interface PersonManageMapper {
    List<PersonManage> getPersons(@Param("name") String name);
    String getMaxidPersons();
    int  savePersons( PersonManage personManage);
    int updatePersons( PersonManage personManage);
    int deletePersons( String  id);

}
