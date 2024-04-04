package com.example.testapachepoi.repository;

import com.example.testapachepoi.entity.Book;
import org.springframework.data.jpa.repository.JpaRepository;

public interface BookRepository extends JpaRepository<Book,Integer> {
}
