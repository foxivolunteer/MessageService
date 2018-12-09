package com.qnb.message.controller;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.qnb.message.entity.User;
import com.qnb.message.repo.UserRepository;

@Controller
@RequestMapping(value = "/")
public class UserController {

	@Autowired
	public UserRepository countryRepo;

	@Autowired
	public Util myUtil;

	@GetMapping()
	public String showPage() {
		return "index";
	}

	@GetMapping("/success")
	public String showSuccess() {
		return "success";
	}

	@GetMapping("/errors")
	public String showError() {
		return "error";
	}

	@PostMapping(value = "/save")
	public String save(User user) throws EncryptedDocumentException, InvalidFormatException, IOException {
		if (user.getName().isEmpty() || user.getSicil().isEmpty() || user.getMessage().isEmpty()) {
			return "redirect:/errors";
		}

		return myUtil.createXLSFile(user);
	}

}
