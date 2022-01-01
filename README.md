# KakaotalkChattingBot_V2

카카오톡 채팅방의 텍스트를 직접 복사하여 긁어오는 방식  
  
  
[v2.2] 대화 연관성 개선
1. 새로운 채팅에 포함되는 채팅 로그를 찾는다. (이 때 chatlog.length >= 5)
2. 찾았으면, 해당 로그 index+1 ~ index+5 에서 가장 유사도가 비슷한 로그를 출력한다.
기준 : Sequence Matcher의 유사도를 이용했다..
