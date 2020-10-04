using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations;

namespace Excel_Interop
{
    class Application_old : DBConnector
    {
        public Application_old(User user)
        {
            Creator = user;
            //Company = user.Company;
            CreatedAt = DateTime.Now;
        }
        
        public User Creator { get; }                   //Создатель заявки
        public DateTime CreatedAt { get; }             //Время создания
        public DateTime EstimatedTime { get; }         //Запланированное время
        public DateTime ClosedAt { get; }              //Время закрытия 
        public TimeSpan ProlongTime { get; }           //Время продления
        public TimeSpan OverdueTime =>
            EstimatedTime > DateTime.Now ?
            DateTime.Now - EstimatedTime :
            TimeSpan.Zero;                      //Просроченное время
        public Classifier Classifier { get; }          //Классификатор заявки
        public User Executor { get; }                  //Исполнитель
        public Company Company { get; }                //УК
        public List<Comment> Comments { get; }         //Комментарии к заявке
        public Address Address { get; }                //Адрес

        
        public void AddComment(User user, string text, byte[] data) =>
            Comments.Add(new Comment(user, text, data, DateTime.Now));
        
        public void RemoveComment(User user, Comment comment) 
        {
            //Query.CommandText = "";

        }
       
        public void ProlongApplication(User user) { }
        
        public void SetExecutor(User user, User executor) { }
       
        public void SetClassifier(User user, Classifier classifier) { }
       
        public void SetAddress(User user, Address address) { }
        
        public void ChangeAddress(User user, Address address) { }
    }
}
