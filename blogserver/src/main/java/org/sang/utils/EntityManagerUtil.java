package org.sang.utils;

import org.hibernate.SQLQuery;
import org.hibernate.transform.Transformers;
import org.springframework.stereotype.Component;
import org.springframework.stereotype.Service;

import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.persistence.Query;
import java.text.Normalizer;
import java.util.List;
import java.util.Map;


@Component
@Service
public  class EntityManagerUtil<T>  {
   @PersistenceContext
   private EntityManager entityManager;


   //1.返回map
   public  List<Map<String, Object>> getListMap(String sql){
      Query nativeQuery=entityManager.createNativeQuery(sql);
      nativeQuery.unwrap(SQLQuery.class).setResultTransformer(Transformers.ALIAS_TO_ENTITY_MAP);
      List resultList=nativeQuery.getResultList();
      return resultList;
   }
  
   //2.返回自定义实体类
   public List<T> nativeQueryResult(String sql, Class clazz) {
     sql = Normalizer.normalize(sql, Normalizer.Form.NFKC);
     sql = sql.replaceAll(".*([';]+|(--)+).*", "");
     Query query = entityManager.createNativeQuery(sql);
     List<T> queryList = query.unwrap(SQLQuery.class).setResultTransformer(Transformers.aliasToBean(clazz)).list();
     return queryList;
   }


}
