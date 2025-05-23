public class CTRL_TestPptxGen {
    
    @AuraEnabled
    public static User getUserData() { 
        return [
            SELECT FirstName, LastName, Username
            FROM User
            WHERE Id =: UserInfo.getUserId()
        ];
    }
}