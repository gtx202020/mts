import oracledb
import logging
from typing import Dict, List, Optional, Tuple, Any

# Oracle Client 경로 설정
ORACLE_CLIENT_PATH = r"C:\instantclient_21_3"

# 로깅 설정
logger = logging.getLogger(__name__)


class DatabaseHandler:
    """데이터베이스 연결 및 메타데이터 조회를 담당하는 클래스"""
    
    def __init__(self):
        """Oracle Client 초기화"""
        try:
            oracledb.init_oracle_client(lib_dir=ORACLE_CLIENT_PATH)
            logger.info("Oracle Client 초기화 성공")
        except Exception as e:
            logger.warning(f"Oracle Client 초기화 실패: {str(e)}")
        
        self.send_connection = None
        self.recv_connection = None
    
    def connect_db(self, sid: str, username: str, password: str, connection_type: str = 'send') -> Optional[oracledb.Connection]:
        """데이터베이스에 연결
        
        Args:
            sid: 데이터베이스 SID
            username: 사용자명
            password: 비밀번호
            connection_type: 'send' 또는 'recv'
        
        Returns:
            DB 연결 객체 또는 None (실패 시)
        """
        try:
            connection = oracledb.connect(user=username, password=password, dsn=sid)
            
            if connection_type == 'send':
                self.send_connection = connection
            else:
                self.recv_connection = connection
                
            logger.info(f"{connection_type} DB 연결 성공: {sid}")
            return connection
            
        except Exception as e:
            logger.error(f"{connection_type} DB 연결 실패: {str(e)}")
            return None
    
    def get_column_info(self, owner: str, table_name: str, connection_type: str = 'send') -> Dict[str, Dict[str, Any]]:
        """테이블의 컬럼 정보 조회
        
        Args:
            owner: 스키마 소유자
            table_name: 테이블명
            connection_type: 'send' 또는 'recv'
        
        Returns:
            컬럼 정보 딕셔너리 {컬럼명: {name, type, size, nullable}}
        """
        connection = self.send_connection if connection_type == 'send' else self.recv_connection
        
        if not connection:
            logger.error(f"{connection_type} DB 연결이 없습니다.")
            return {}
        
        try:
            cursor = connection.cursor()
            query = """
                SELECT column_name, data_type, data_length, nullable
                FROM all_tab_columns
                WHERE owner = :owner
                AND table_name = :table_name
                ORDER BY column_id
            """
            
            cursor.execute(query, owner=owner, table_name=table_name)
            columns = {}
            
            for row in cursor:
                columns[row[0]] = {
                    'name': row[0],
                    'type': row[1],
                    'size': str(row[2]),
                    'nullable': 'Y' if row[3] == 'Y' else 'N'
                }
            
            cursor.close()
            logger.info(f"{owner}.{table_name} 테이블에서 {len(columns)}개 컬럼 정보 조회 완료")
            return columns
            
        except Exception as e:
            logger.error(f"컬럼 정보 조회 실패: {str(e)}")
            return {}
    
    def close_connections(self):
        """모든 DB 연결 종료"""
        if self.send_connection:
            try:
                self.send_connection.close()
                logger.info("송신 DB 연결 종료")
            except:
                pass
                
        if self.recv_connection:
            try:
                self.recv_connection.close()
                logger.info("수신 DB 연결 종료")
            except:
                pass


class ColumnValidator:
    """컬럼 검증 로직을 담당하는 클래스"""
    
    def __init__(self):
        """초기화"""
        self.varchar_types = ['VARCHAR', 'VARCHAR2', 'CHAR']
        
    def check_column_exists(self, column_name: str, columns_info: Dict[str, Dict[str, Any]], 
                          column_type: str) -> Tuple[bool, Optional[str]]:
        """컬럼 존재 여부 검사
        
        Args:
            column_name: 검사할 컬럼명
            columns_info: 테이블의 컬럼 정보
            column_type: '송신' 또는 '수신'
        
        Returns:
            (존재여부, 에러메시지)
        """
        if not column_name:
            return True, None  # 빈 컬럼은 정상으로 처리
            
        if column_name not in columns_info:
            return False, f"{column_type} 테이블에 {column_name} 컬럼이 존재하지 않습니다."
            
        return True, None
    
    def check_type_compatibility(self, send_type: str, recv_type: str) -> Tuple[bool, Optional[str]]:
        """타입 호환성 검사
        
        Args:
            send_type: 송신 컬럼 타입
            recv_type: 수신 컬럼 타입
        
        Returns:
            (호환여부, 경고메시지)
        """
        # 동일 타입
        if send_type == recv_type:
            return True, None
            
        # VARCHAR 계열 호환
        if send_type in self.varchar_types and recv_type in self.varchar_types:
            return True, None
            
        # 다른 타입
        return False, f"타입이 다릅니다: 송신({send_type}) vs 수신({recv_type})"
    
    def check_date_varchar_conversion(self, send_type: str, recv_type: str) -> Optional[str]:
        """DATE-VARCHAR 변환 감지
        
        Args:
            send_type: 송신 컬럼 타입
            recv_type: 수신 컬럼 타입
        
        Returns:
            경고 메시지 또는 None
        """
        # DATE → VARCHAR 변환
        if send_type == 'DATE' and recv_type in self.varchar_types:
            return "DATE → VARCHAR 변환: 날짜 형식 확인 필요"
            
        # VARCHAR → DATE 변환
        if send_type in self.varchar_types and recv_type == 'DATE':
            return "VARCHAR → DATE 변환: 날짜 형식 검증 필요"
            
        return None
    
    def check_size_compatibility(self, send_info: Dict[str, Any], recv_info: Dict[str, Any]) -> Optional[str]:
        """크기 호환성 검사
        
        Args:
            send_info: 송신 컬럼 정보
            recv_info: 수신 컬럼 정보
        
        Returns:
            경고 메시지 또는 None
        """
        send_type = send_info.get('type', '')
        recv_type = recv_info.get('type', '')
        
        # DATE 타입은 크기 비교 제외
        if 'DATE' in (send_type, recv_type):
            return None
            
        # VARCHAR 계열만 크기 비교
        if send_type in self.varchar_types and recv_type in self.varchar_types:
            try:
                send_size = int(send_info.get('size', '0'))
                recv_size = int(recv_info.get('size', '0'))
                
                if send_size > recv_size:
                    return f"크기 불일치: 송신({send_size}) > 수신({recv_size})"
                    
            except ValueError:
                return "크기 정보 변환 오류"
                
        return None
    
    def check_nullable_compatibility(self, send_nullable: str, recv_nullable: str) -> Optional[str]:
        """NULL 허용 여부 호환성 검사
        
        Args:
            send_nullable: 송신 컬럼 NULL 허용 여부 ('Y' 또는 'N')
            recv_nullable: 수신 컬럼 NULL 허용 여부 ('Y' 또는 'N')
        
        Returns:
            경고 메시지 또는 None
        """
        if send_nullable == 'Y' and recv_nullable == 'N':
            return "NULL 제약조건 위반 가능: 송신(NULL 허용) → 수신(NOT NULL)"
            
        return None
    
    def validate_columns(self, send_mapping: List[str], recv_mapping: List[str],
                        send_columns: Dict[str, Dict[str, Any]], 
                        recv_columns: Dict[str, Dict[str, Any]]) -> List[Dict[str, Any]]:
        """전체 컬럼 검증 수행
        
        Args:
            send_mapping: 송신 컬럼 매핑 리스트
            recv_mapping: 수신 컬럼 매핑 리스트
            send_columns: 송신 테이블 컬럼 정보
            recv_columns: 수신 테이블 컬럼 정보
        
        Returns:
            검증 결과 리스트
        """
        results = []
        
        # 매핑 리스트 길이 맞추기
        max_len = max(len(send_mapping), len(recv_mapping))
        send_mapping = send_mapping + [''] * (max_len - len(send_mapping))
        recv_mapping = recv_mapping + [''] * (max_len - len(recv_mapping))
        
        for idx, (send_col, recv_col) in enumerate(zip(send_mapping, recv_mapping)):
            result = {
                'send_column': send_col,
                'recv_column': recv_col,
                'send_info': None,
                'recv_info': None,
                'errors': [],
                'warnings': [],
                'status': '정상'
            }
            
            # 송신 컬럼 검사
            if send_col:
                exists, error = self.check_column_exists(send_col, send_columns, '송신')
                if not exists:
                    result['errors'].append(error)
                else:
                    result['send_info'] = send_columns[send_col]
            
            # 수신 컬럼 검사
            if recv_col:
                exists, error = self.check_column_exists(recv_col, recv_columns, '수신')
                if not exists:
                    result['errors'].append(error)
                else:
                    result['recv_info'] = recv_columns[recv_col]
            
            # 둘 다 존재하는 경우 추가 검증
            if result['send_info'] and result['recv_info']:
                send_info = result['send_info']
                recv_info = result['recv_info']
                
                # 타입 호환성 검사
                compatible, warning = self.check_type_compatibility(
                    send_info['type'], recv_info['type']
                )
                if not compatible:
                    result['warnings'].append(warning)
                
                # DATE-VARCHAR 변환 검사
                date_warning = self.check_date_varchar_conversion(
                    send_info['type'], recv_info['type']
                )
                if date_warning:
                    result['warnings'].append(date_warning)
                
                # 크기 호환성 검사
                size_warning = self.check_size_compatibility(send_info, recv_info)
                if size_warning:
                    result['warnings'].append(size_warning)
                
                # NULL 허용 여부 검사
                null_warning = self.check_nullable_compatibility(
                    send_info['nullable'], recv_info['nullable']
                )
                if null_warning:
                    result['warnings'].append(null_warning)
            
            # 상태 결정
            if result['errors']:
                result['status'] = '오류'
            elif result['warnings']:
                result['status'] = '경고'
            
            results.append(result)
        
        return results