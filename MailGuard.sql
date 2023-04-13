# Creation script - adds MGCheck storedprocedure and required tables in hmaildb. 
# Execute this on your hmail server database


SET FOREIGN_KEY_CHECKS = 0;

CREATE DEFINER=`root`@`%` FUNCTION `MGCheck`(_fromEmail varchar(50),_toEmail Varchar(5000)) RETURNS varchar(5000) CHARSET utf8
BEGIN
-- MailGuard MySQL embedded version 
-- Get the CSV of allowed IDs for the fromEmail value - store in allowedAddresses
-- Loop through the toEmail addresses, check if the exact address or *@domain are in allowedAddresses. If not found, add the address to Result
 

 Set session group_concat_max_len=100000;
 
 Set @Result="";
 Set @allowedList = (Select IFNULL(group_Concat(ToEmail),'') from mgRulesTable where Enabled="y" and fromEmail=_fromEmail);

	/* if @allowedList = "" THEN 
			Return "X";
	End if;
	*/
  IF (Select Instr(@allowedList,"*@*")<>0) THEN 
			RETURN "";
	END IF;
	
  WHILE _toEmail != '' DO
		-- Grab each email ID in the list
    SET @element = SUBSTRING_INDEX(_toEmail, ',', 1);      
				
		If (Select count(*) from mgallowLocalDomainsTable where fromEmail=_fromEmail)=1 THEN #This sender can send to any address on any local domain
			#Now: If the recipient's domain is a local domain, allow delivery
			If (Select count(*) from hm_domains where domainName=reverse(SUBSTRING_INDEX(reverse(@element), '@', 1)))<> 1 THEN # Temporarily disabled 
			-- Check for both *@domain and exact match of toEmail in the allowed list. If both are not found, add the address to list of blocked addresses
				IF (Select Instr(@allowedList,Concat("*@",reverse(SUBSTRING_INDEX(reverse(@element), '@', 1)))) + Instr(@allowedList,@element)) = 0 THEN
						Set @Result = Concat(@element,",",@Result);
				END IF;
			END IF;
		ELSE #non-privileged user, check for exact match
				-- Check for both *@domain and exact match of toEmail in the allowed list. If both are not found, add the address to list of blocked addresses
				IF (Select Instr(@allowedList,Concat("*@",reverse(SUBSTRING_INDEX(reverse(@element), '@', 1)))) + Instr(@allowedList,@element)) = 0 THEN
						Set @Result = Concat(@element,",",@Result);
				END IF;
		END IF;
		
    IF LOCATE(',', _toEmail) > 0 THEN
      SET _toEmail = SUBSTRING(_toEmail, LOCATE(',', _toEmail) + 1);
    ELSE
      SET _toEmail = '';
    END IF;

  END WHILE;

Return @Result;

END

-- ----------------------------
-- Table structure for mgallowlocaldomainstable
-- ----------------------------
DROP TABLE IF EXISTS `mgallowlocaldomainstable`;
CREATE TABLE `mgallowlocaldomainstable`  (
  `FromEmail` varchar(255) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  PRIMARY KEY (`FromEmail`) USING BTREE
) ENGINE = InnoDB CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Dynamic;

-- ----------------------------
-- Table structure for mgrulestable
-- ----------------------------
DROP TABLE IF EXISTS `mgrulestable`;
CREATE TABLE `mgrulestable`  (
  `RuleId` int(255) NOT NULL AUTO_INCREMENT,
  `FromEmail` varchar(255) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `ToEmail` varchar(255) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `Remarks` varchar(200) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT '',
  `LastModified` timestamp NOT NULL DEFAULT current_timestamp() ON UPDATE CURRENT_TIMESTAMP,
  `Enabled` char(1) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT 'y' COMMENT 'y/n - Is this rule enabled or not',
  PRIMARY KEY (`RuleId`) USING BTREE,
  UNIQUE INDEX `c`(`FromEmail`, `ToEmail`) USING BTREE,
  INDEX `a`(`FromEmail`) USING BTREE,
  INDEX `b`(`ToEmail`) USING BTREE
) ENGINE = InnoDB AUTO_INCREMENT = 4614 CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Dynamic;

-- ----------------------------
-- Table structure for mgviolationstable
-- ----------------------------
DROP TABLE IF EXISTS `mgviolationstable`;
CREATE TABLE `mgviolationstable`  (
  `Sno` int(255) NOT NULL AUTO_INCREMENT,
  `FromEmail` varchar(255) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL,
  `FromIP` varchar(18) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT NULL,
  `ToEmail` varchar(255) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL DEFAULT '',
  `Message` varchar(200) CHARACTER SET utf8 COLLATE utf8_general_ci NULL DEFAULT '',
  `EventTime` timestamp NOT NULL DEFAULT current_timestamp(),
  PRIMARY KEY (`Sno`) USING BTREE
) ENGINE = InnoDB AUTO_INCREMENT = 1 CHARACTER SET = utf8 COLLATE = utf8_general_ci ROW_FORMAT = Dynamic;

SET FOREIGN_KEY_CHECKS = 1;
